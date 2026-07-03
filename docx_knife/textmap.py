"""Reversible visible-text map for a single ``<w:p>`` element.

Phase-3 building block for paragraph-internal editing. The public helpers here
are *internal* to :mod:`docx_knife` — Phase 6 wires them into the batch
operation layer but never re-exports them.

Design references: openspec/changes/docx-patch-engine/design.md §6, §4.2 and
spec `text-map-editing`.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any

from lxml import etree

from ._models import Selector
from .errors import InvalidContentError, InvalidPatternError, UnsupportedStructureError

# Namespace-qualified tag names, kept module-local to avoid cluttering the
# shared _ooxml surface with symbols only Phase-3 editing consumes.
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{_W_NS}}}{local}"


_T = _q("t")
_R = _q("r")
_TAB = _q("tab")
_BR = _q("br")
_CR = _q("cr")
_INS = _q("ins")
_DEL = _q("del")
_INSTR_TEXT = _q("instrText")
_SDT = _q("sdt")
_SDT_CONTENT = _q("sdtContent")
_HYPERLINK = _q("hyperlink")
_FLD_SIMPLE = _q("fldSimple")
_FLD_CHAR = _q("fldChar")
_BOOKMARK_START = _q("bookmarkStart")
_BOOKMARK_END = _q("bookmarkEnd")
_PERM_START = _q("permStart")
_PERM_END = _q("permEnd")
_COMMENT_START = _q("commentRangeStart")
_COMMENT_END = _q("commentRangeEnd")
_W_TYPE = _q("type")

# Ancestor tags that "colour" the visible characters they contain. Point-markers
# (bookmarkStart/End etc.) are not ancestors of characters; we ignore them for
# range intersection because their capability policy is "preserve".
_TRACKED_ANCESTORS: frozenset[str] = frozenset({
    _HYPERLINK,
    _FLD_SIMPLE,
    _INS,
    _DEL,
    _INSTR_TEXT,
})

# Capability matrix — see design §4.2 / spec `text-map-editing`.
_CAPABILITY_MATRIX: dict[str, str] = {
    _HYPERLINK: "reject",
    _FLD_SIMPLE: "reject",
    _FLD_CHAR: "reject",
    _INSTR_TEXT: "reject",
    _BOOKMARK_START: "preserve",
    _BOOKMARK_END: "preserve",
    _PERM_START: "reject",
    _PERM_END: "reject",
    _COMMENT_START: "preserve",
    _COMMENT_END: "preserve",
    _INS: "allow",
    _DEL: "reject",
}

# Human-readable structure name used in error metadata and structures_in_range.
_STRUCT_LABELS: dict[str, str] = {
    _HYPERLINK: "w:hyperlink",
    _FLD_SIMPLE: "w:fldSimple",
    _FLD_CHAR: "w:fldChar",
    _INSTR_TEXT: "w:instrText",
    _BOOKMARK_START: "w:bookmarkStart",
    _BOOKMARK_END: "w:bookmarkEnd",
    _PERM_START: "w:permStart",
    _PERM_END: "w:permEnd",
    _COMMENT_START: "w:commentRangeStart",
    _COMMENT_END: "w:commentRangeEnd",
    _INS: "w:ins",
    _DEL: "w:del",
}

_MARKER_TO_LITERAL: dict[str, str] = {
    "TAB": "[[DOCX:TAB]]",
    "LINE_BREAK": "[[DOCX:LINE_BREAK]]",
    "PAGE_BREAK": "[[DOCX:PAGE_BREAK]]",
    "COLUMN_BREAK": "[[DOCX:COLUMN_BREAK]]",
    "CR": "[[DOCX:CR]]",
}

_MARKER_NAMES: frozenset[str] = frozenset(_MARKER_TO_LITERAL)


@dataclass(frozen=True, slots=True)
class TextPosition:
    """One character position in :class:`TextMap`.

    ``marker`` is ``None`` for characters that came from a ``<w:t>``. For
    projected non-text nodes (tab, break, cr) it names the marker kind; every
    character of one marker's projected string shares the same ``node_ref`` and
    ``marker``.
    """

    node_ref: Any
    node_offset: int
    text_offset: int
    run_ref: Any | None
    marker: str | None


@dataclass(frozen=True, slots=True)
class TextMap:
    text: str
    positions: tuple[TextPosition, ...]
    atomic_ranges: tuple[tuple[int, int], ...]
    intersected_structures_at: Mapping[int, tuple[str, ...]]
    # Ordinal (within the paragraph, document order) for each contributing
    # node (``<w:t>`` or marker element). Consumers use it to compute
    # ``node_range`` / ``crosses_nodes`` without re-walking the tree.
    node_ordinals: Mapping[int, int]

    def structures_in_range(self, start: int, end: int) -> tuple[str, ...]:
        """Return stable-ordered qualified structure tags intersecting ``[start, end)``."""
        if start >= end:
            return ()
        seen: dict[str, None] = {}
        for offset in range(start, end):
            for tag in self.intersected_structures_at.get(offset, ()):
                seen.setdefault(tag, None)
        return tuple(sorted(seen))


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------


def _br_marker_kind(br_elem: etree._Element) -> str:
    kind = br_elem.get(_W_TYPE)
    if kind == "page":
        return "PAGE_BREAK"
    if kind == "column":
        return "COLUMN_BREAK"
    return "LINE_BREAK"


def _nearest_run(elem: etree._Element, paragraph: etree._Element) -> etree._Element | None:
    cur: etree._Element | None = elem
    while cur is not None and cur is not paragraph:
        if cur.tag == _R:
            return cur
        cur = cur.getparent()
    return None


def _escape_literal_docx(text: str) -> tuple[str, list[int], list[tuple[int, int]]]:
    """Escape literal ``[[DOCX:`` occurrences with a leading backslash.

    Returns the escaped string, per-character source offsets (a synthetic
    backslash reuses the offset of the '[' it precedes), and the *relative*
    span of each escape sequence within the escaped string. Callers translate
    these relative spans into global atomic ranges.
    """
    out_chars: list[str] = []
    origin: list[int] = []
    escape_spans: list[tuple[int, int]] = []
    i = 0
    n = len(text)
    marker = "[[DOCX:"
    mlen = len(marker)
    while i < n:
        if text[i] == "[" and text.startswith(marker, i):
            span_start = len(out_chars)
            out_chars.append("\\")
            origin.append(i)
            for j in range(mlen):
                out_chars.append(text[i + j])
                origin.append(i + j)
            escape_spans.append((span_start, len(out_chars)))
            i += mlen
        else:
            out_chars.append(text[i])
            origin.append(i)
            i += 1
    return "".join(out_chars), origin, escape_spans


def build_text_map(paragraph: etree._Element) -> TextMap:
    """Build the visible-text ``TextMap`` for one paragraph.

    The walk descends in document order, includes ``<w:ins>`` descendants and
    excludes ``<w:del>`` and ``<w:instrText>`` descendants (final Word view).
    Adjacent ``<w:t>`` elements form one continuous text string.
    """
    text_parts: list[str] = []
    positions: list[TextPosition] = []
    atomic_ranges: list[tuple[int, int]] = []
    structures_at: dict[int, tuple[str, ...]] = {}
    node_ordinals: dict[int, int] = {}
    ordinal_counter = 0

    active_structures: list[str] = []
    global_offset = 0

    def record_char(
        char: str,
        node_ref: etree._Element,
        node_offset: int,
        run_ref: etree._Element | None,
        marker: str | None,
    ) -> None:
        nonlocal global_offset
        text_parts.append(char)
        positions.append(
            TextPosition(
                node_ref=node_ref,
                node_offset=node_offset,
                text_offset=global_offset,
                run_ref=run_ref,
                marker=marker,
            )
        )
        if active_structures:
            structures_at[global_offset] = tuple(active_structures)
        global_offset += 1

    def visit(elem: etree._Element) -> None:
        nonlocal ordinal_counter
        tag = elem.tag
        # Skip whole subtrees that are not part of the final visible view.
        if tag in (_DEL, _INSTR_TEXT, _SDT, _SDT_CONTENT):
            return

        pushed_struct = False
        struct_label: str | None = None
        if tag in _TRACKED_ANCESTORS:
            struct_label = _STRUCT_LABELS.get(tag)
            if struct_label is not None:
                active_structures.append(struct_label)
                pushed_struct = True

        try:
            if tag == _T:
                raw = elem.text or ""
                escaped, origin, escape_spans = _escape_literal_docx(raw)
                run_ref = _nearest_run(elem, paragraph)
                if escaped:
                    node_ordinals[id(elem)] = ordinal_counter
                    ordinal_counter += 1
                    base_offset = global_offset
                    for ch, src in zip(escaped, origin, strict=True):
                        record_char(ch, elem, src, run_ref, None)
                    for rel_start, rel_end in escape_spans:
                        atomic_ranges.append((base_offset + rel_start, base_offset + rel_end))
                return
            if tag == _TAB:
                literal = _MARKER_TO_LITERAL["TAB"]
                run_ref = _nearest_run(elem, paragraph)
                node_ordinals[id(elem)] = ordinal_counter
                ordinal_counter += 1
                start = global_offset
                for i, ch in enumerate(literal):
                    record_char(ch, elem, i, run_ref, "TAB")
                atomic_ranges.append((start, global_offset))
                return
            if tag == _BR:
                kind = _br_marker_kind(elem)
                literal = _MARKER_TO_LITERAL[kind]
                run_ref = _nearest_run(elem, paragraph)
                node_ordinals[id(elem)] = ordinal_counter
                ordinal_counter += 1
                start = global_offset
                for i, ch in enumerate(literal):
                    record_char(ch, elem, i, run_ref, kind)
                atomic_ranges.append((start, global_offset))
                return
            if tag == _CR:
                literal = _MARKER_TO_LITERAL["CR"]
                run_ref = _nearest_run(elem, paragraph)
                node_ordinals[id(elem)] = ordinal_counter
                ordinal_counter += 1
                start = global_offset
                for i, ch in enumerate(literal):
                    record_char(ch, elem, i, run_ref, "CR")
                atomic_ranges.append((start, global_offset))
                return

            # Point-markers (bookmarkStart/permEnd/commentRangeStart/…): do not
            # contribute characters, but we still remember which offsets they
            # sit between via the ``intersected_structures_at`` map at the
            # *next* character emitted (they surround it).
            if tag in _STRUCT_LABELS and tag not in _TRACKED_ANCESTORS:
                # Nothing to emit; children (if any) walk normally below.
                pass

            for child in elem:
                visit(child)
        finally:
            if pushed_struct:
                active_structures.pop()

    for child in paragraph:
        visit(child)

    return TextMap(
        text="".join(text_parts),
        positions=tuple(positions),
        atomic_ranges=tuple(atomic_ranges),
        intersected_structures_at=structures_at,
        node_ordinals=dict(node_ordinals),
    )


# ---------------------------------------------------------------------------
# Marker restoration for writeback
# ---------------------------------------------------------------------------


def restore_markers(text: str) -> list[tuple[str, str]]:
    """Split ``text`` into ``('text', literal)`` / ``('marker', name)`` segments.

    Escaped literal ``\\[[DOCX:`` sequences un-escape to ``[[DOCX:``.
    Any ``[[DOCX:...]]`` token that is not a known marker name raises
    :class:`InvalidContentError`.
    """
    segments: list[tuple[str, str]] = []
    buf: list[str] = []
    i = 0
    n = len(text)
    escape_prefix = "\\[[DOCX:"
    marker_prefix = "[[DOCX:"

    def flush() -> None:
        if buf:
            segments.append(("text", "".join(buf)))
            buf.clear()

    while i < n:
        if text.startswith(escape_prefix, i):
            buf.append(marker_prefix)
            i += len(escape_prefix)
            continue
        if text.startswith(marker_prefix, i):
            end = text.find("]]", i + len(marker_prefix))
            if end == -1:
                raise InvalidContentError(
                    raw=False,
                    reason=f"unterminated reserved marker at offset {i}",
                )
            name = text[i + len(marker_prefix) : end]
            if name not in _MARKER_NAMES:
                raise InvalidContentError(
                    raw=False,
                    reason=f"unknown reserved marker: [[DOCX:{name}]]",
                )
            flush()
            segments.append(("marker", name))
            i = end + 2
            continue
        buf.append(text[i])
        i += 1
    flush()
    return segments


# ---------------------------------------------------------------------------
# Capability evaluation
# ---------------------------------------------------------------------------


def evaluate_capability(
    intersected: tuple[str, ...],
    char_range: tuple[int, int],
    *,
    target_id: str = "",
) -> None:
    """Raise :class:`UnsupportedStructureError` when any structure rejects the range."""
    reject: list[str] = []
    reverse = {label: tag for tag, label in _STRUCT_LABELS.items()}
    for label in intersected:
        tag = reverse.get(label)
        if tag is None:
            continue
        if _CAPABILITY_MATRIX.get(tag) == "reject":
            reject.append(label)
    if reject:
        raise UnsupportedStructureError(
            target_id=target_id,
            structures=tuple(sorted(set(reject))),
            matched_range=char_range,
        )


def range_hits_atomic(
    atomic_ranges: tuple[tuple[int, int], ...],
    start: int,
    end: int,
) -> bool:
    """True iff a boundary of ``[start, end)`` falls strictly inside a marker."""
    for a, b in atomic_ranges:
        if a < start < b:
            return True
        if a < end < b:
            return True
    return False


# ---------------------------------------------------------------------------
# Selectors
# ---------------------------------------------------------------------------


SelectorMatcher = Callable[[str], list[tuple[int, int]]]


def compile_selector(selector: Selector) -> SelectorMatcher:
    """Compile a :class:`Selector` into a callable that returns non-overlapping matches."""
    if selector.regex:
        try:
            compiled = re.compile(selector.pattern, re.DOTALL)
        except re.error as exc:
            raise InvalidPatternError(pattern=selector.pattern, reason=str(exc)) from exc

        def _find_regex(haystack: str) -> list[tuple[int, int]]:
            out: list[tuple[int, int]] = []
            pos = 0
            while pos <= len(haystack):
                m = compiled.search(haystack, pos)
                if m is None:
                    break
                s, e = m.start(), m.end()
                if e == s:
                    # Skip zero-length matches: advance one char and retry.
                    pos = s + 1
                    continue
                out.append((s, e))
                pos = e
            return out

        return _find_regex

    if not selector.pattern:
        raise InvalidPatternError(
            pattern=selector.pattern, reason="literal pattern must not be empty"
        )
    needle = selector.pattern

    def _find_literal(haystack: str) -> list[tuple[int, int]]:
        out: list[tuple[int, int]] = []
        start = 0
        while True:
            idx = haystack.find(needle, start)
            if idx < 0:
                break
            end = idx + len(needle)
            out.append((idx, end))
            start = end
        return out

    return _find_literal


def select_matches(
    matches: list[tuple[int, int]],
    occurrence: int | None,
) -> list[tuple[int, int]]:
    """Resolve occurrence semantics.

    Callers (paragraph.py) raise :class:`AmbiguousTextMatchError` /
    :class:`TextNotFoundError` when this helper returns an empty list or the
    ``ambiguous`` sentinel because they hold the ``target_id`` needed by the
    error. This helper only distinguishes the resolution *shape*.
    """
    total = len(matches)
    if occurrence is None:
        if total == 0:
            return []
        if total == 1:
            return [matches[0]]
        # Ambiguous — signal via sentinel: caller checks total via ``matches``.
        return matches  # caller must detect ambiguity by len > 1 & occurrence is None
    if occurrence == -1:
        # right-to-left ordering keeps earlier character positions valid
        return list(reversed(matches))
    if 0 <= occurrence < total:
        return [matches[occurrence]]
    return []


__all__ = [
    "TextMap",
    "TextPosition",
    "build_text_map",
    "compile_selector",
    "evaluate_capability",
    "range_hits_atomic",
    "restore_markers",
    "select_matches",
]
