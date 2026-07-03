"""Paragraph-internal text editors (Phase 3).

Public consumers get a fluent :class:`Paragraph` object in Phase 4; this module
implements only the *mutators* that operate on a single ``<w:p>`` element in
place. Phase 6 wires them into the batch layer and holds the ``target_id``
for structured errors.
"""

from __future__ import annotations

import copy
import weakref
from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from ._models import ContentItem, Selector
from .errors import (
    AmbiguousTextMatchError,
    InvalidContentError,
    ParagraphNotFoundError,
    TextNotFoundError,
    UnsupportedStructureError,
)
from .textmap import (
    TextMap,
    TextPosition,
    build_text_map,
    compile_selector,
    evaluate_capability,
    range_hits_atomic,
    restore_markers,
)

if TYPE_CHECKING:
    from .document import Document

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _q(local: str) -> str:
    return f"{{{_W_NS}}}{local}"


_R = _q("r")
_T = _q("t")
_RPR = _q("rPr")
_TAB = _q("tab")
_BR = _q("br")
_CR = _q("cr")
_TYPE = _q("type")

_XML_SPACE = f"{{{_XML_NS}}}space"

_MARKER_TAGS: dict[str, str] = {
    "TAB": _TAB,
    "LINE_BREAK": _BR,
    "PAGE_BREAK": _BR,
    "COLUMN_BREAK": _BR,
    "CR": _CR,
}

_MARKER_ATTR: dict[str, str | None] = {
    "TAB": None,
    "LINE_BREAK": None,
    "PAGE_BREAK": "page",
    "COLUMN_BREAK": "column",
    "CR": None,
}

_PREVIEW_LIMIT = 80


@dataclass(frozen=True, slots=True)
class ReplaceOutcome:
    before_preview: str
    after_preview: str
    warnings: tuple[str, ...] = ()


def _preview(text: str) -> str:
    if len(text) <= _PREVIEW_LIMIT:
        return text
    return text[: _PREVIEW_LIMIT - 1] + "\u2026"


# ---------------------------------------------------------------------------
# Run construction
# ---------------------------------------------------------------------------


def _clone_run_shell(run: etree._Element) -> etree._Element:
    """Return a new empty ``<w:r>`` carrying a deep copy of ``run``'s ``<w:rPr>``."""
    new_run = etree.Element(_R)
    rpr = run.find(_RPR)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    return new_run


def _make_run(rpr: etree._Element | None) -> etree._Element:
    run = etree.Element(_R)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    return run


def _append_text_child(run: etree._Element, text: str) -> None:
    if not text:
        return
    t = etree.SubElement(run, _T)
    if text[:1].isspace() or text[-1:].isspace():
        t.set(_XML_SPACE, "preserve")
    t.text = text


def _append_marker_child(run: etree._Element, marker: str) -> None:
    tag = _MARKER_TAGS[marker]
    node = etree.SubElement(run, tag)
    attr_val = _MARKER_ATTR[marker]
    if attr_val is not None:
        node.set(_TYPE, attr_val)


def _build_replacement_run(
    text: str, rpr_template: etree._Element | None
) -> etree._Element | None:
    """Build a single ``<w:r>`` for ``text``, or ``None`` when ``text`` is empty.

    Reserved markers become dedicated child nodes (``<w:tab/>``,
    ``<w:br/>``, ``<w:cr/>``) while literal text is placed in one or more
    ``<w:t>`` children. All children share the same ``<w:rPr>``.
    """
    segments = restore_markers(text)
    if not segments:
        return None
    run = _make_run(rpr_template)
    for kind, payload in segments:
        if kind == "text":
            _append_text_child(run, payload)
        else:
            _append_marker_child(run, payload)
    return run


# ---------------------------------------------------------------------------
# DOM helpers
# ---------------------------------------------------------------------------


def _child_index(parent: etree._Element, child: etree._Element) -> int:
    for i, sibling in enumerate(parent):
        if sibling is child:
            return i
    raise ValueError("child not in parent")


def _run_has_structural_content(run: etree._Element) -> bool:
    for child in run:
        tag = child.tag
        if tag == _RPR:
            continue
        if tag == _T:
            continue
        return True
    return False


def _cleanup_run(run: etree._Element) -> None:
    """Drop empty ``<w:t>`` children and remove the run if only ``<w:rPr>`` remains.

    Runs that still host drawings/tabs/breaks are preserved untouched.
    """
    if _run_has_structural_content(run):
        return
    for t in list(run):
        if t.tag == _T and (t.text is None or t.text == ""):
            run.remove(t)
    remaining = [c for c in run if c.tag != _RPR]
    if not remaining:
        parent = run.getparent()
        if parent is not None:
            parent.remove(run)


# ---------------------------------------------------------------------------
# Splice / insert
# ---------------------------------------------------------------------------


def _split_run_before(
    run: etree._Element,
    child: etree._Element,
    text_offset: int | None,
) -> tuple[etree._Element, etree._Element]:
    """Split ``run`` at the boundary immediately before ``child`` (or inside it).

    * ``text_offset is None`` — split just before ``child``. ``child`` and all
      later siblings move to the new right run.
    * ``text_offset > 0`` — ``child`` must be a ``<w:t>``; its text is sliced
      so the left side keeps ``text[:text_offset]`` and the right side gets a
      new ``<w:t>`` holding ``text[text_offset:]`` followed by ``child``'s
      trailing siblings.
    * ``text_offset == 0`` — equivalent to splitting immediately before ``child``.

    Returns ``(left_run, right_run)``. ``left_run`` is the mutated original run;
    ``right_run`` is a freshly-inserted sibling immediately after it.
    """
    parent = run.getparent()
    assert parent is not None
    right_run = _clone_run_shell(run)

    children = list(run)
    child_idx = children.index(child)
    trailing = children[child_idx + 1 :]

    if text_offset is None or text_offset == 0:
        run.remove(child)
        right_run.append(child)
    else:
        assert child.tag == _T, "text_offset requires a <w:t> child"
        original_text = child.text or ""
        left_text = original_text[:text_offset]
        right_text = original_text[text_offset:]
        child.text = left_text or None
        if left_text and (left_text[:1].isspace() or left_text[-1:].isspace()):
            child.set(_XML_SPACE, "preserve")
        if right_text:
            new_t = etree.SubElement(right_run, _T)
            if right_text[:1].isspace() or right_text[-1:].isspace():
                new_t.set(_XML_SPACE, "preserve")
            new_t.text = right_text
    for sibling in trailing:
        run.remove(sibling)
        right_run.append(sibling)

    run_index = _child_index(parent, run)
    parent.insert(run_index + 1, right_run)
    return run, right_run


def _split_run_after(
    run: etree._Element,
    child: etree._Element,
    text_offset: int | None,
) -> tuple[etree._Element, etree._Element]:
    """Split ``run`` at the boundary immediately after ``child`` (or inside it).

    * ``text_offset is None`` — split right after ``child``; ``child`` stays in
      the left run and its trailing siblings move to the right run.
    * ``text_offset >= 0`` — ``child`` must be a ``<w:t>``; the left side keeps
      ``text[:text_offset+1]`` and the right side gets ``text[text_offset+1:]``.
    """
    if text_offset is None:
        parent = run.getparent()
        assert parent is not None
        right_run = _clone_run_shell(run)
        children = list(run)
        child_idx = children.index(child)
        trailing = children[child_idx + 1 :]
        for sibling in trailing:
            run.remove(sibling)
            right_run.append(sibling)
        run_index = _child_index(parent, run)
        parent.insert(run_index + 1, right_run)
        return run, right_run

    # child must be a <w:t> here.
    return _split_run_before(run, child, text_offset + 1)


def _splice_match(
    positions: tuple[TextPosition, ...],
    replacement_run: etree._Element | None,
) -> None:
    """Excise ``positions`` from the paragraph and inject ``replacement_run``.

    Split the paragraph so the matched span occupies one contiguous slice of
    runs, remove those runs, then insert the replacement in the gap.
    """
    if not positions:
        return
    first_pos = positions[0]
    last_pos = positions[-1]
    first_node: etree._Element = first_pos.node_ref
    last_node: etree._Element = last_pos.node_ref

    first_run = first_node.getparent()
    last_run = last_node.getparent()
    assert first_run is not None
    assert last_run is not None
    run_parent = first_run.getparent()
    assert run_parent is not None

    # Single-node fast paths avoid the trap that _split_run_before creates a
    # *new* <w:t> for the right half, invalidating our stale last_node.
    if first_node is last_node:
        if first_node.tag == _T:
            first_offset = min(p.node_offset for p in positions)
            last_offset = max(p.node_offset for p in positions)
            original = first_node.text or ""
            left = original[:first_offset]
            right = original[last_offset + 1 :]
            # Update host <w:t> to the left prefix.
            first_node.text = left or None
            if left and (left[:1].isspace() or left[-1:].isspace()):
                first_node.set(_XML_SPACE, "preserve")
            # Move the trailing siblings after first_node into a new right run
            # so the replacement can be inserted between them and first_run.
            children = list(first_run)
            pos = children.index(first_node)
            trailing = children[pos + 1 :]
            first_index = _child_index(run_parent, first_run)
            insert_at = first_index + 1

            if right or trailing:
                right_run = _clone_run_shell(first_run)
                if right:
                    _append_text_child(right_run, right)
                for sibling in trailing:
                    first_run.remove(sibling)
                    right_run.append(sibling)
                run_parent.insert(insert_at, right_run)
                if replacement_run is not None:
                    run_parent.insert(insert_at, replacement_run)
            elif replacement_run is not None:
                run_parent.insert(insert_at, replacement_run)
            _cleanup_run(first_run)
            return
        # Single marker consumed.
        parent = first_node.getparent()
        insert_at = _child_index(run_parent, first_run) + 1
        # Preserve trailing siblings by splitting them into a new run.
        run_children = list(first_run)
        pos = run_children.index(first_node)
        trailing = run_children[pos + 1 :]
        if parent is not None:
            parent.remove(first_node)
        if trailing:
            right_run = _clone_run_shell(first_run)
            for sibling in trailing:
                first_run.remove(sibling)
                right_run.append(sibling)
            run_parent.insert(insert_at, right_run)
        if replacement_run is not None:
            run_parent.insert(insert_at, replacement_run)
        _cleanup_run(first_run)
        return

    # Multi-node case: split off the head, then split off the tail, and remove
    # the runs sandwiched between them.
    if first_node.tag == _T:
        first_offset = min(p.node_offset for p in positions if p.node_ref is first_node)
        _, first_right = _split_run_before(first_run, first_node, first_offset)
    else:
        _, first_right = _split_run_before(first_run, first_node, None)

    if last_run is first_run:
        last_run = first_right

    if last_node.tag == _T:
        last_offset = max(p.node_offset for p in positions if p.node_ref is last_node)
        _, last_right = _split_run_after(last_run, last_node, last_offset)
    else:
        _, last_right = _split_run_after(last_run, last_node, None)

    to_remove: list[etree._Element] = []
    node: etree._Element | None = first_right
    while node is not None:
        to_remove.append(node)
        if node is last_run:
            break
        node = node.getnext()

    insertion_index = _child_index(run_parent, first_right)
    for r in to_remove:
        run_parent.remove(r)

    if replacement_run is not None:
        run_parent.insert(insertion_index, replacement_run)

    _cleanup_run(first_run)
    _cleanup_run(last_right)


def _insert_at_boundary(
    text_map: TextMap,
    boundary_position: TextPosition,
    boundary_is_left: bool,
    replacement_run: etree._Element | None,
) -> None:
    """Insert ``replacement_run`` immediately before or after ``boundary_position``.

    ``boundary_is_left`` True inserts *before* the boundary character; False
    inserts *after* it.
    """
    if replacement_run is None:
        return
    host_node: etree._Element = boundary_position.node_ref
    host_run = host_node.getparent()
    assert host_run is not None
    run_parent = host_run.getparent()
    assert run_parent is not None

    if host_node.tag == _T:
        if boundary_is_left:
            _, right_run = _split_run_before(
                host_run, host_node, boundary_position.node_offset
            )
        else:
            _, right_run = _split_run_after(
                host_run, host_node, boundary_position.node_offset
            )
    else:
        # Marker: never split inside; only around.
        if boundary_is_left:
            _, right_run = _split_run_before(host_run, host_node, None)
        else:
            _, right_run = _split_run_after(host_run, host_node, None)

    right_index = _child_index(run_parent, right_run)
    run_parent.insert(right_index, replacement_run)
    _cleanup_run(host_run)
    _cleanup_run(right_run)


# ---------------------------------------------------------------------------
# Precondition checks per match
# ---------------------------------------------------------------------------


def _check_match_capability(
    text_map: TextMap, start: int, end: int, *, target_id: str
) -> None:
    structs = text_map.structures_in_range(start, end)
    evaluate_capability(structs, (start, end), target_id=target_id)
    if range_hits_atomic(text_map.atomic_ranges, start, end):
        raise UnsupportedStructureError(
            target_id=target_id,
            structures=("reserved_marker",),
            matched_range=(start, end),
        )


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------


def _resolve_matches(
    paragraph: etree._Element,
    selector: Selector,
    occurrence: int | None,
    *,
    target_id: str,
) -> tuple[TextMap, list[tuple[int, int]]]:
    text_map = build_text_map(paragraph)
    matcher = compile_selector(selector)
    matches = matcher(text_map.text)
    total = len(matches)
    if occurrence is None:
        if total == 0:
            raise TextNotFoundError(
                target_id=target_id,
                selector=selector,
                occurrence=None,
                total_matches=0,
            )
        if total > 1:
            raise AmbiguousTextMatchError(
                target_id=target_id,
                selector=selector,
                total_matches=total,
            )
        return text_map, [matches[0]]
    if occurrence == -1:
        if total == 0:
            raise TextNotFoundError(
                target_id=target_id,
                selector=selector,
                occurrence=-1,
                total_matches=0,
            )
        # Right-to-left keeps earlier offsets valid without a rebuild.
        return text_map, list(reversed(matches))
    if occurrence < 0 or occurrence >= total:
        raise TextNotFoundError(
            target_id=target_id,
            selector=selector,
            occurrence=occurrence,
            total_matches=total,
        )
    return text_map, [matches[occurrence]]


def _reject_raw(raw: bool) -> None:
    if raw:
        raise InvalidContentError(
            raw=True,
            reason="paragraph-internal operations do not accept raw=true",
        )


def _rpr_for_position(position: TextPosition) -> etree._Element | None:
    run = position.run_ref
    if run is None:
        return None
    found: etree._Element | None = run.find(_RPR)
    return found


def apply_replace_text(
    paragraph: etree._Element,
    find: Selector,
    replacement: str,
    occurrence: int | None,
    *,
    target_id: str,
    raw: bool = False,
) -> ReplaceOutcome:
    _reject_raw(raw)
    # Validate replacement's marker syntax up front (fail fast before mutating).
    restore_markers(replacement)
    text_map, matches = _resolve_matches(paragraph, find, occurrence, target_id=target_id)
    before = _preview(text_map.text)
    for i, (start, end) in enumerate(matches):
        if i > 0:
            text_map = build_text_map(paragraph)
            # Match offsets shift once earlier edits mutate the map when we
            # iterate right-to-left; but sequentially rebuilding is safer
            # against overlaps.
        _check_match_capability(text_map, start, end, target_id=target_id)
        positions = text_map.positions[start:end]
        rpr_template = _rpr_for_position(positions[0])
        replacement_run = _build_replacement_run(replacement, rpr_template)
        _splice_match(positions, replacement_run)
    after = _preview(build_text_map(paragraph).text)
    return ReplaceOutcome(before_preview=before, after_preview=after)


def apply_delete_text(
    paragraph: etree._Element,
    find: Selector,
    occurrence: int | None,
    *,
    target_id: str,
    raw: bool = False,
) -> ReplaceOutcome:
    _reject_raw(raw)
    text_map, matches = _resolve_matches(paragraph, find, occurrence, target_id=target_id)
    before = _preview(text_map.text)
    for i, (start, end) in enumerate(matches):
        if i > 0:
            text_map = build_text_map(paragraph)
        _check_match_capability(text_map, start, end, target_id=target_id)
        positions = text_map.positions[start:end]
        _splice_match(positions, None)
    after = _preview(build_text_map(paragraph).text)
    return ReplaceOutcome(before_preview=before, after_preview=after)


def apply_insert_text_before(
    paragraph: etree._Element,
    find: Selector,
    insertion: str,
    occurrence: int | None,
    *,
    target_id: str,
    raw: bool = False,
) -> ReplaceOutcome:
    _reject_raw(raw)
    restore_markers(insertion)
    text_map, matches = _resolve_matches(paragraph, find, occurrence, target_id=target_id)
    before = _preview(text_map.text)
    for i, (start, end) in enumerate(matches):
        if i > 0:
            text_map = build_text_map(paragraph)
        _check_match_capability(text_map, start, end, target_id=target_id)
        boundary = text_map.positions[start]
        rpr_template = _rpr_for_position(boundary)
        replacement_run = _build_replacement_run(insertion, rpr_template)
        _insert_at_boundary(
            text_map,
            boundary,
            boundary_is_left=True,
            replacement_run=replacement_run,
        )
    after = _preview(build_text_map(paragraph).text)
    return ReplaceOutcome(before_preview=before, after_preview=after)


def apply_insert_text_after(
    paragraph: etree._Element,
    find: Selector,
    insertion: str,
    occurrence: int | None,
    *,
    target_id: str,
    raw: bool = False,
) -> ReplaceOutcome:
    _reject_raw(raw)
    restore_markers(insertion)
    text_map, matches = _resolve_matches(paragraph, find, occurrence, target_id=target_id)
    before = _preview(text_map.text)
    for i, (start, end) in enumerate(matches):
        if i > 0:
            text_map = build_text_map(paragraph)
        _check_match_capability(text_map, start, end, target_id=target_id)
        boundary = text_map.positions[end - 1]
        rpr_template = _rpr_for_position(boundary)
        replacement_run = _build_replacement_run(insertion, rpr_template)
        _insert_at_boundary(
            text_map,
            boundary,
            boundary_is_left=False,
            replacement_run=replacement_run,
        )
    after = _preview(build_text_map(paragraph).text)
    return ReplaceOutcome(before_preview=before, after_preview=after)


__all__ = [
    "Paragraph",
    "ReplaceOutcome",
    "apply_delete_text",
    "apply_insert_text_after",
    "apply_insert_text_before",
    "apply_replace_text",
]


# ---------------------------------------------------------------------------
# Fluent Paragraph object
# ---------------------------------------------------------------------------


class Paragraph:
    """Fluent handle to a live paragraph inside a :class:`Document`.

    Every method re-resolves the underlying ``<w:p>`` node through the
    document's anchor manifest. If the ID has been invalidated (for example by
    ``replace_para``), the next call raises :class:`ParagraphNotFoundError`.
    """

    __slots__ = ("id", "_doc_ref")

    def __init__(self, doc: Document, target_id: str) -> None:
        self.id = target_id
        self._doc_ref = weakref.ref(doc)

    # ---------------------------------------------------------------- helpers

    def _doc(self) -> Document:
        doc = self._doc_ref()
        if doc is None:
            raise ParagraphNotFoundError(target_id=self.id)
        return doc

    def _element(self) -> etree._Element:
        return self._doc()._manifest.resolve(self.id)

    def __repr__(self) -> str:
        return f"Paragraph(id={self.id!r})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Paragraph):
            return NotImplemented
        return self.id == other.id and self._doc_ref() is other._doc_ref()

    def __hash__(self) -> int:
        return hash((self.id, id(self._doc_ref())))

    # ------------------------------------------------------------------ read

    def read(self, *, raw: bool = False) -> str:
        return self._doc().get_paragraph(self.id, raw=raw)

    # -------------------------------------------------------------- paragraph

    def insert_para_before(
        self,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        return self._doc().insert_para_before(
            self.id, items, raw=raw, normalize_text=normalize_text
        )

    def insert_para_after(
        self,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        return self._doc().insert_para_after(
            self.id, items, raw=raw, normalize_text=normalize_text
        )

    def replace_para(
        self,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        return self._doc().replace_para(
            self.id, items, raw=raw, normalize_text=normalize_text
        )

    def delete_para(self) -> None:
        self._doc().delete_para([self.id])

    # --------------------------------------------------------- text (in-place)

    def replace_text(
        self,
        find: str | Selector,
        replacement: str,
        *,
        occurrence: int | None = None,
        normalize_text: bool = False,
    ) -> None:
        element = self._element()
        selector = Selector.coerce(find)
        text = _maybe_normalize(replacement, normalize_text)
        apply_replace_text(
            element, selector, text, occurrence, target_id=self.id
        )

    def delete_text(
        self,
        find: str | Selector,
        *,
        occurrence: int | None = None,
    ) -> None:
        element = self._element()
        selector = Selector.coerce(find)
        apply_delete_text(element, selector, occurrence, target_id=self.id)

    def insert_text_before(
        self,
        find: str | Selector,
        text: str,
        *,
        occurrence: int | None = None,
        normalize_text: bool = False,
    ) -> None:
        element = self._element()
        selector = Selector.coerce(find)
        payload = _maybe_normalize(text, normalize_text)
        apply_insert_text_before(
            element, selector, payload, occurrence, target_id=self.id
        )

    def insert_text_after(
        self,
        find: str | Selector,
        text: str,
        *,
        occurrence: int | None = None,
        normalize_text: bool = False,
    ) -> None:
        element = self._element()
        selector = Selector.coerce(find)
        payload = _maybe_normalize(text, normalize_text)
        apply_insert_text_after(
            element, selector, payload, occurrence, target_id=self.id
        )


def _maybe_normalize(text: str, enabled: bool) -> str:
    if not enabled:
        return text
    from .content import normalize

    return normalize(text)
