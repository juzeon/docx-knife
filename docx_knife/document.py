"""Document lifecycle and Phase-2 read/query APIs.

Persistence (``save``) is provided by the sibling :mod:`docx_knife.save`
module; the ``Document`` class only exposes the public surface here.
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
import tempfile
import zipfile
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from types import TracebackType

from lxml import etree

from . import _ooxml
from ._models import (
    AnyEditOperation,
    ContentItem,
    EditResult,
    Pagination,
    ParagraphInfo,
    ParagraphListResult,
    ParagraphLocation,
    ParagraphMatchInfo,
    ParagraphSearchResult,
    SaveResult,
    Selector,
    TextMatch,
)
from ._paragraph_ops import (
    build_new_paragraph,
    detect_protected_structures,
    expand_visible_items,
    parse_raw_paragraphs,
)
from .anchors import AnchorManifest
from .content import ContentResolverConfig, normalize, resolve_items
from .errors import (
    AmbiguousTextMatchError,
    DocumentNotFoundError,
    InvalidContentError,
    InvalidDocumentError,
    InvalidPatternError,
    ParagraphNotFoundError,
    TextNotFoundError,
    ValidationError,
)
from .paragraph import Paragraph

MAIN_PART: str = "word/document.xml"


@dataclass(frozen=True, slots=True)
class _SourceFingerprint:
    sha256: str
    size: int
    mtime_ns: int


@dataclass(slots=True)
class _ParagraphRecord:
    target_id: str
    node: etree._Element
    location: ParagraphLocation
    style_id: str | None
    w14_para_id: str | None


@dataclass(slots=True)
class _Index:
    records: tuple[_ParagraphRecord, ...] = ()
    id_to_record: dict[str, _ParagraphRecord] = field(default_factory=dict)


class Document:
    """A DOCX opened for querying and editing via the docx-knife patch engine."""

    def __init__(
        self,
        *,
        source_path: Path,
        temp_dir: Path,
        tree: etree._ElementTree,
        fingerprint: _SourceFingerprint,
        content_config: ContentResolverConfig | None = None,
    ) -> None:
        self._source_path = source_path
        self._temp_dir = temp_dir
        self._tree = tree
        self._root = tree.getroot()
        self._source_fingerprint = fingerprint
        self._manifest = AnchorManifest(self._root)
        self._index_cache: _Index | None = None
        self._closed = False
        self._content_config: ContentResolverConfig = (
            content_config
            if content_config is not None
            else ContentResolverConfig(
                workspace_root=temp_dir,
                input_roots=(source_path.parent.resolve(),),
            )
        )
        self._last_op_warnings: tuple[str, ...] = ()
        self._change_log: list[dict[str, object]] = []
        self._build_index()

    # ------------------------------------------------------------------ open

    @classmethod
    def open(
        cls,
        source_path: str | os.PathLike[str],
        *,
        content_config: ContentResolverConfig | None = None,
    ) -> Document:
        """Open ``source_path`` as a DOCX and return a live :class:`Document`."""
        path = Path(source_path)
        if not path.is_file():
            raise DocumentNotFoundError(path=str(path))
        raw_bytes = path.read_bytes()
        fingerprint = _fingerprint(path, raw_bytes)
        temp_dir = Path(tempfile.mkdtemp(prefix="docx-knife-"))
        try:
            tree = _load_main_document(path, temp_dir)
        except BaseException:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise
        return cls(
            source_path=path,
            temp_dir=temp_dir,
            tree=tree,
            fingerprint=fingerprint,
            content_config=content_config,
        )

    # ----------------------------------------------------------- lifecycle

    def close(self) -> None:
        """Release the private temp workspace. Idempotent."""
        if self._closed:
            return
        self._closed = True
        shutil.rmtree(self._temp_dir, ignore_errors=True)

    def __enter__(self) -> Document:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()

    def save(self, output_path: str | os.PathLike[str]) -> SaveResult:
        """Save the current DOM to ``output_path`` with source-drift check and ``.bak`` backup."""
        if self._closed:
            raise RuntimeError("document is closed")
        from .save import save_document

        return save_document(self, output_path)

    # --------------------------------------------------------------- state

    @property
    def source_path(self) -> Path:
        return self._source_path

    @property
    def _fingerprint(self) -> _SourceFingerprint:
        return self._source_fingerprint

    # -------------------------------------------------------------- index

    def _build_index(self) -> None:
        records: list[_ParagraphRecord] = []
        id_to_record: dict[str, _ParagraphRecord] = {}
        for node, location, style_id, w14_para_id in _ooxml.iter_editable_paragraphs(self._root):
            existing = self._manifest.id_for_node(node)
            target_id = existing if existing is not None else self._manifest.allocate()
            if existing is None:
                self._manifest.bind(target_id, node)
            record = _ParagraphRecord(
                target_id=target_id,
                node=node,
                location=location,
                style_id=style_id,
                w14_para_id=w14_para_id,
            )
            records.append(record)
            id_to_record[target_id] = record
        self._index_cache = _Index(records=tuple(records), id_to_record=id_to_record)

    def _invalidate_index(self) -> None:
        self._index_cache = None

    def _ensure_index(self) -> _Index:
        if self._index_cache is None:
            self._build_index()
        assert self._index_cache is not None
        return self._index_cache

    # -------------------------------------------------------------- reads

    def paragraph_count(self) -> int:
        """Return the number of editable paragraphs currently in the document."""
        return len(self._ensure_index().records)

    def list_paragraphs(
        self,
        start: int = 1,
        limit: int | None = None,
        max_chars: int = 80,
        raw: bool = False,
    ) -> ParagraphListResult:
        """Return a paginated slice of paragraphs with truncated previews."""
        records = self._ensure_index().records
        total = len(records)
        window = _window_slice(records, start=start, limit=limit)
        infos = tuple(_record_to_info(record, max_chars=max_chars, raw=raw) for record in window)
        return ParagraphListResult(
            paragraphs=infos,
            pagination=Pagination(start=start, limit=limit, returned=len(infos), total=total),
        )

    def get_paragraph(self, paragraph_id: str, raw: bool = False) -> str:
        """Return the full text (or raw XML) of one paragraph."""
        record = self._get_record(paragraph_id)
        return _record_content(record, raw=raw)

    def get_visible_text(self, raw: bool = False) -> str:
        """Return the concatenated visible text of every paragraph, in document order."""
        records = self._ensure_index().records
        if raw:
            return "".join(_ooxml.serialize_paragraph(rec.node) for rec in records)
        return "\n".join(_ooxml.visible_text_plain(rec.node) for rec in records)

    # ------------------------------------------------------------- search

    def grep_paragraphs(
        self,
        pattern: str,
        regex: bool = False,
        start: int = 1,
        limit: int | None = None,
        max_chars: int = 0,
        raw: bool = False,
    ) -> ParagraphSearchResult:
        """Search paragraphs by literal or regex ``pattern`` and return hits with ranges."""
        matcher = _compile_matcher(pattern, regex=regex)
        records = self._ensure_index().records
        window = _window_slice(records, start=start, limit=limit)
        matches: list[ParagraphMatchInfo] = []
        total_matches = 0
        for record in window:
            haystack = _record_content(record, raw=raw)
            ranges = matcher.find_all(haystack)
            if not ranges:
                continue
            total_matches += len(ranges)
            info = _record_to_info(record, max_chars=max_chars, raw=raw)
            matches.append(
                ParagraphMatchInfo(
                    paragraph=info,
                    ranges=tuple(ranges),
                    match_count=len(ranges),
                )
            )
        return ParagraphSearchResult(
            matches=tuple(matches),
            total_matches=total_matches,
            pagination=Pagination(
                start=start,
                limit=limit,
                returned=len(matches),
                total=len(records),
            ),
        )

    def count_matches(
        self,
        pattern: str,
        regex: bool = False,
        paragraph_id: str | None = None,
        raw: bool = False,
    ) -> int:
        """Return the total number of matches for ``pattern``.

        If ``paragraph_id`` is set, count within that paragraph only.
        """
        matcher = _compile_matcher(pattern, regex=regex)
        total = 0
        for record in self._search_records(paragraph_id):
            total += len(matcher.find_all(_record_content(record, raw=raw)))
        return total

    def find_text(
        self,
        pattern: str,
        regex: bool = False,
        occurrence: int | None = None,
        paragraph_id: str | None = None,
        raw: bool = False,
    ) -> TextMatch | list[TextMatch] | None:
        """Locate matches for ``pattern`` and return a ``TextMatch`` (single, list, or ``None``)."""
        matcher = _compile_matcher(pattern, regex=regex)
        all_matches: list[TextMatch] = []
        for record in self._search_records(paragraph_id):
            haystack = _record_content(record, raw=raw)
            ranges = matcher.find_all(haystack)
            if not ranges:
                continue
            if raw:
                ordinals: list[int] = []
            else:
                ordinals = _ooxml.visible_char_node_ordinals(record.node)
            for char_range in ranges:
                start, end = char_range
                if raw or not ordinals:
                    node_range = char_range
                    crosses = False
                else:
                    start_ordinal = ordinals[start] if start < len(ordinals) else ordinals[-1]
                    end_ordinal = ordinals[end - 1] if 0 < end <= len(ordinals) else ordinals[-1]
                    node_range = (start_ordinal, end_ordinal)
                    crosses = start_ordinal != end_ordinal
                all_matches.append(
                    TextMatch(
                        paragraph_id=record.target_id,
                        char_range=char_range,
                        node_range=node_range,
                        crosses_nodes=crosses,
                        total_matches=0,
                    )
                )
        total = len(all_matches)
        all_matches = [
            TextMatch(
                paragraph_id=m.paragraph_id,
                char_range=m.char_range,
                node_range=m.node_range,
                crosses_nodes=m.crosses_nodes,
                total_matches=total,
            )
            for m in all_matches
        ]

        if occurrence is None:
            if total == 0:
                return None
            if total > 1:
                raise AmbiguousTextMatchError(
                    target_id=paragraph_id or "*",
                    selector=Selector(pattern=pattern, regex=regex),
                    total_matches=total,
                )
            return all_matches[0]
        if occurrence == -1:
            return list(all_matches)
        if 0 <= occurrence < total:
            return all_matches[occurrence]
        raise TextNotFoundError(
            target_id=paragraph_id or "*",
            selector=Selector(pattern=pattern, regex=regex),
            occurrence=occurrence,
            total_matches=total,
        )

    # -------------------------------------------------------- paragraph write

    def insert_para_before(
        self,
        target_id: str,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        """Insert ``items`` immediately before the paragraph identified by
        ``target_id``. Returns the new paragraphs in document order.

        Style is inherited from the previous-sibling paragraph when one exists
        so that inserting before a heading continues the preceding body flow
        rather than duplicating the heading's formatting. When ``target_id``
        has no previous sibling paragraph (e.g. it opens a section), the
        target itself is used as the style anchor.
        """
        target = self._manifest.resolve(target_id)
        previous = _previous_paragraph_sibling(target)
        style_anchor = previous if previous is not None else target
        new_elements = self._expand_items_to_elements(
            anchor=style_anchor, items=items, raw=raw, normalize_text=normalize_text
        )
        parent = target.getparent()
        assert parent is not None
        insert_index = parent.index(target)
        for element in new_elements:
            parent.insert(insert_index, element)
            insert_index += 1
        return self._register_new_paragraphs(new_elements)

    def insert_para_after(
        self,
        target_id: str,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        """Insert ``items`` immediately after the paragraph identified by
        ``target_id``. Uses a moving-cursor algorithm to preserve item order."""
        anchor = self._manifest.resolve(target_id)
        new_elements = self._expand_items_to_elements(
            anchor=anchor, items=items, raw=raw, normalize_text=normalize_text
        )
        cursor = anchor
        for element in new_elements:
            cursor.addnext(element)
            cursor = element
        return self._register_new_paragraphs(new_elements)

    def replace_para(
        self,
        target_id: str,
        items: Sequence[str | ContentItem],
        *,
        raw: bool = False,
        normalize_text: bool = False,
    ) -> list[Paragraph]:
        """Replace the paragraph identified by ``target_id`` with ``items``.

        Runs a wholesale replacement even when it would strip protected
        structures; every detected structure is recorded in ``warnings`` on
        the document (readable via ``last_operation_warnings``)."""
        target = self._manifest.resolve(target_id)
        warnings = detect_protected_structures(target)
        new_elements = self._expand_items_to_elements(
            anchor=target, items=items, raw=raw, normalize_text=normalize_text
        )
        parent = target.getparent()
        assert parent is not None
        insert_index = parent.index(target)
        for element in new_elements:
            parent.insert(insert_index, element)
            insert_index += 1
        parent.remove(target)
        self._manifest.invalidate(target_id)
        result = self._register_new_paragraphs(new_elements)
        self._last_op_warnings = warnings
        return result

    def delete_para(self, target_ids: Sequence[str]) -> None:
        """Delete the paragraphs identified by ``target_ids``.

        Fails fast when the collection is empty, contains duplicates, or names
        an unknown ID. Nodes are removed in reverse document order."""
        ids = list(target_ids)
        checks = ("nonempty", "unique", "resolvable")
        if not ids:
            raise ValidationError(stage="prevalidation", checks=checks, failed_check="nonempty")
        if len(set(ids)) != len(ids):
            raise ValidationError(stage="prevalidation", checks=checks, failed_check="unique")
        resolved: list[etree._Element] = [self._manifest.resolve(tid) for tid in ids]

        # Order by document position (identity match against iter_editable_paragraphs).
        position_by_id: dict[int, int] = {}
        for pos, (node, _loc, _style, _paraid) in enumerate(
            _ooxml.iter_editable_paragraphs(self._root)
        ):
            position_by_id[id(node)] = pos
        pairs = sorted(
            zip(ids, resolved, strict=True),
            key=lambda pair: position_by_id[id(pair[1])],
            reverse=True,
        )
        for tid, element in pairs:
            parent = element.getparent()
            assert parent is not None
            parent.remove(element)
            self._manifest.invalidate(tid)
        self._invalidate_index()
        self._last_op_warnings = ()

    def get_paragraph_object(self, paragraph_id: str) -> Paragraph:
        """Return a fluent :class:`Paragraph` handle bound to ``paragraph_id``."""
        # Force validation of the ID before handing out a handle.
        self._get_record(paragraph_id)
        return Paragraph(self, paragraph_id)

    def batch_edit(
        self,
        operations: Sequence[AnyEditOperation],
        *,
        normalize_text: bool = False,
        envelope: dict[str, object] | None = None,
    ) -> EditResult:
        """Execute ``operations`` atomically. See :mod:`docx_knife.batch`."""
        from .batch import BatchExecutor

        return BatchExecutor(
            self,
            operations,
            normalize_text=normalize_text,
            envelope=envelope,
        ).run()

    def change_log(self) -> list[dict[str, object]]:
        """Return a shallow copy of the current change-log entries."""
        return list(self._change_log)

    @property
    def last_operation_warnings(self) -> tuple[str, ...]:
        """Warnings recorded by the most recent write operation."""
        return self._last_op_warnings

    # ------------------------------------------------------ paragraph helpers

    def _expand_items_to_elements(
        self,
        *,
        anchor: etree._Element,
        items: Sequence[str | ContentItem],
        raw: bool,
        normalize_text: bool,
    ) -> list[etree._Element]:
        coerced = tuple(ContentItem.of(item) for item in items)
        if not coerced:
            raise InvalidContentError(raw=raw, reason="items must be non-empty")
        resolved = resolve_items(coerced, raw=raw, config=self._content_config)
        if raw:
            new_elements: list[etree._Element] = []
            for item in resolved:
                fragment = item.paragraphs[0]
                new_elements.extend(parse_raw_paragraphs(fragment))
            return new_elements
        texts = expand_visible_items(resolved)
        if normalize_text:
            texts = [normalize(text) for text in texts]
        return [build_new_paragraph(anchor, text) for text in texts]

    def _register_new_paragraphs(self, elements: Sequence[etree._Element]) -> list[Paragraph]:
        handles: list[Paragraph] = []
        for element in elements:
            new_id = self._manifest.allocate()
            self._manifest.bind(new_id, element)
            handles.append(Paragraph(self, new_id))
        self._invalidate_index()
        return handles

    # ------------------------------------------------------------- helpers

    def _get_record(self, paragraph_id: str) -> _ParagraphRecord:
        record = self._ensure_index().id_to_record.get(paragraph_id)
        if record is None:
            raise ParagraphNotFoundError(target_id=paragraph_id)
        # Validate the manifest still binds this ID to the same live node.
        self._manifest.resolve(paragraph_id)
        return record

    def _search_records(self, paragraph_id: str | None) -> tuple[_ParagraphRecord, ...]:
        if paragraph_id is None:
            return self._ensure_index().records
        return (self._get_record(paragraph_id),)


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------


def _previous_paragraph_sibling(element: etree._Element) -> etree._Element | None:
    """Return the closest preceding ``<w:p>`` sibling of ``element``, or ``None``.

    Non-paragraph siblings (``<w:sectPr>``, ``<w:tbl>``, …) are skipped so the
    lookup stays scoped to the surrounding body flow.
    """
    sibling = element.getprevious()
    while sibling is not None:
        if sibling.tag == _ooxml.P_TAG:
            return sibling
        sibling = sibling.getprevious()
    return None


def _fingerprint(path: Path, data: bytes) -> _SourceFingerprint:
    stat = path.stat()
    digest = hashlib.sha256(data).hexdigest()
    return _SourceFingerprint(sha256=digest, size=stat.st_size, mtime_ns=stat.st_mtime_ns)


def _load_main_document(source_path: Path, temp_dir: Path) -> etree._ElementTree:
    path_str = str(source_path)
    try:
        with zipfile.ZipFile(source_path) as zf:
            if MAIN_PART not in zf.namelist():
                raise InvalidDocumentError(
                    path=path_str,
                    reason=f"missing required part {MAIN_PART!r}",
                )
            zf.extractall(temp_dir)
    except zipfile.BadZipFile as exc:
        raise InvalidDocumentError(path=path_str, reason="not a valid ZIP archive") from exc
    main_path = temp_dir / MAIN_PART
    parser = _ooxml.build_secure_parser()
    try:
        tree = etree.parse(str(main_path), parser)
    except etree.XMLSyntaxError as exc:
        raise InvalidDocumentError(
            path=path_str,
            reason=f"cannot parse {MAIN_PART}: {exc.msg}",
        ) from exc
    return tree


def _window_slice(
    records: tuple[_ParagraphRecord, ...],
    *,
    start: int,
    limit: int | None,
) -> tuple[_ParagraphRecord, ...]:
    if start < 1:
        raise ValueError("start must be >= 1")
    if limit is not None and limit < 0:
        raise ValueError("limit must be >= 0 when provided")
    begin = start - 1
    if begin >= len(records):
        return ()
    if limit is None:
        return records[begin:]
    return records[begin : begin + limit]


def _record_content(record: _ParagraphRecord, *, raw: bool) -> str:
    if raw:
        return _ooxml.serialize_paragraph(record.node)
    return _ooxml.visible_text_plain(record.node)


def _record_to_info(record: _ParagraphRecord, *, max_chars: int, raw: bool) -> ParagraphInfo:
    content = _record_content(record, raw=raw)
    truncated = _truncate(content, max_chars)
    text = None if raw else truncated
    xml = truncated if raw else None
    return ParagraphInfo(
        id=record.target_id,
        global_ordinal=record.location.global_ordinal,
        style_id=record.style_id,
        location=record.location,
        text=text,
        xml=xml,
        w14_para_id=record.w14_para_id,
    )


def _truncate(value: str, max_chars: int) -> str:
    if max_chars <= 0 or len(value) <= max_chars:
        return value
    return value[:max_chars]


class _Matcher:
    def __init__(self, finder: Callable[[str], list[tuple[int, int]]]) -> None:
        self._finder = finder

    def find_all(self, haystack: str) -> list[tuple[int, int]]:
        return self._finder(haystack)


def _compile_matcher(pattern: str, *, regex: bool) -> _Matcher:
    if regex:
        try:
            compiled = re.compile(pattern)
        except re.error as exc:
            raise InvalidPatternError(pattern=pattern, reason=str(exc)) from exc

        def _find_regex(haystack: str) -> list[tuple[int, int]]:
            return [(m.start(), m.end()) for m in compiled.finditer(haystack)]

        return _Matcher(_find_regex)

    if not pattern:
        raise InvalidPatternError(pattern=pattern, reason="literal pattern must not be empty")
    needle = pattern

    def _find_literal(haystack: str) -> list[tuple[int, int]]:
        found: list[tuple[int, int]] = []
        start = 0
        while True:
            idx = haystack.find(needle, start)
            if idx < 0:
                break
            end = idx + len(needle)
            found.append((idx, end))
            start = end
        return found

    return _Matcher(_find_literal)


__all__ = ["Document"]
