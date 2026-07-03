"""Immutable public models for ``docx_knife``.

All types are frozen dataclasses or ``Literal`` unions. Downstream phases will
extend the operation hierarchy; the shape defined here is the contract the
JSON schema and the Python API both bind to.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from typing import Any, ClassVar, Literal, NewType, TypeAlias

from .errors import InvalidContentError

ParagraphId = NewType("ParagraphId", str)

VerticalMerge = Literal["none", "restart", "continue"]
PartName = Literal["word/document.xml"]


@dataclass(frozen=True, slots=True)
class TableContext:
    """Logical and physical position of a paragraph inside a table cell."""

    table_index: int
    row_index: int
    physical_cell_index: int
    logical_column_index: int
    grid_span: int
    grid_before: int
    vertical_merge: VerticalMerge
    nesting_depth: int
    paragraph_index_in_cell: int


@dataclass(frozen=True, slots=True)
class ParagraphLocation:
    """Structural location of a paragraph within ``word/document.xml``."""

    part: PartName
    original_index: int
    global_ordinal: int
    table_context: TableContext | None = None


@dataclass(frozen=True, slots=True)
class ParagraphInfo:
    """Summary record for a paragraph returned by a query API."""

    id: str
    global_ordinal: int
    style_id: str | None
    location: ParagraphLocation
    text: str | None = None
    xml: str | None = None
    w14_para_id: str | None = None

    def __post_init__(self) -> None:
        if (self.text is None) == (self.xml is None):
            raise InvalidContentError(
                raw=self.xml is not None,
                reason="ParagraphInfo requires exactly one of text or xml",
            )


@dataclass(frozen=True, slots=True)
class Pagination:
    """Window metadata attached to paginated query results."""

    start: int
    limit: int | None
    returned: int
    total: int


@dataclass(frozen=True, slots=True)
class ParagraphListResult:
    """Paginated result of ``Document.list_paragraphs``."""

    paragraphs: tuple[ParagraphInfo, ...]
    pagination: Pagination


@dataclass(frozen=True, slots=True)
class ParagraphMatchInfo:
    """Paragraph summary bundled with the character ranges that matched."""

    paragraph: ParagraphInfo
    ranges: tuple[tuple[int, int], ...]
    match_count: int


@dataclass(frozen=True, slots=True)
class ParagraphSearchResult:
    """Paginated result of ``Document.grep_paragraphs``."""

    matches: tuple[ParagraphMatchInfo, ...]
    total_matches: int
    pagination: Pagination


@dataclass(frozen=True, slots=True)
class TextMatch:
    """A resolved ``find_text`` hit anchored to a paragraph ID."""

    paragraph_id: str
    char_range: tuple[int, int]
    node_range: tuple[int, int]
    crosses_nodes: bool
    total_matches: int
    intersected_structures: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class Selector:
    """Literal-or-regex text selector shared by all text APIs."""

    pattern: str
    regex: bool = False

    @classmethod
    def coerce(cls, value: str | Selector | Mapping[str, Any]) -> Selector:
        if isinstance(value, Selector):
            return value
        if isinstance(value, str):
            return cls(pattern=value, regex=False)
        if isinstance(value, Mapping):
            try:
                pattern = value["pattern"]
            except KeyError as exc:
                raise InvalidContentError(
                    raw=False,
                    reason="selector mapping missing required 'pattern'",
                ) from exc
            if not isinstance(pattern, str):
                raise InvalidContentError(
                    raw=False,
                    reason="selector 'pattern' must be a string",
                )
            regex = value.get("regex", False)
            if not isinstance(regex, bool):
                raise InvalidContentError(
                    raw=False,
                    reason="selector 'regex' must be a boolean",
                )
            return cls(pattern=pattern, regex=regex)
        raise InvalidContentError(
            raw=False,
            reason=f"cannot coerce {type(value).__name__} into a Selector",
        )


@dataclass(frozen=True, slots=True)
class ContentSourceJsonPath:
    """JSONPath reference to a single scalar in a JSON file."""

    source: str
    path: str
    type: Literal["jsonpath"] = "jsonpath"


@dataclass(frozen=True, slots=True)
class ContentSourceFile:
    """UTF-8 text file loaded from an allowed input root."""

    path: str
    encoding: str = "utf-8"
    type: Literal["file"] = "file"


@dataclass(frozen=True, slots=True)
class ContentSourceCommand:
    """External command whose stdout supplies the item's text."""

    argv: tuple[str, ...]
    timeout_seconds: float = 30.0
    env: Mapping[str, str] | None = None
    cwd: str | None = None
    type: Literal["command"] = "command"


ContentRef: TypeAlias = ContentSourceJsonPath | ContentSourceFile | ContentSourceCommand


@dataclass(frozen=True, slots=True)
class ContentItem:
    """One unit of content: either an inline literal or a reference."""

    content_literal: str | None = None
    content_ref: ContentRef | None = None

    def __post_init__(self) -> None:
        if (self.content_literal is None) == (self.content_ref is None):
            raise InvalidContentError(
                raw=False,
                reason="ContentItem requires exactly one of content_literal or content_ref",
            )

    @classmethod
    def of(cls, value: str | ContentItem | Mapping[str, Any]) -> ContentItem:
        if isinstance(value, ContentItem):
            return value
        if isinstance(value, str):
            return cls(content_literal=value)
        if isinstance(value, Mapping):
            has_literal = "content_literal" in value
            has_ref = "content_ref" in value
            if has_literal == has_ref:
                raise InvalidContentError(
                    raw=False,
                    reason=(
                        "ContentItem mapping requires exactly one of content_literal or content_ref"
                    ),
                )
            if has_literal:
                literal = value["content_literal"]
                if not isinstance(literal, str):
                    raise InvalidContentError(
                        raw=False,
                        reason="content_literal must be a string",
                    )
                return cls(content_literal=literal)
            return cls(content_ref=_coerce_content_ref(value["content_ref"]))
        raise InvalidContentError(
            raw=False,
            reason=f"cannot coerce {type(value).__name__} into a ContentItem",
        )


def _coerce_content_ref(value: Any) -> ContentRef:
    if isinstance(value, (ContentSourceJsonPath, ContentSourceFile, ContentSourceCommand)):
        return value
    if not isinstance(value, Mapping):
        raise InvalidContentError(
            raw=False,
            reason="content_ref must be a mapping",
        )
    kind = value.get("type")
    if kind == "jsonpath":
        return ContentSourceJsonPath(
            source=_required_str(value, "source"),
            path=_required_str(value, "path"),
        )
    if kind == "file":
        return ContentSourceFile(
            path=_required_str(value, "path"),
            encoding=str(value.get("encoding", "utf-8")),
        )
    if kind == "command":
        argv = value.get("argv")
        if not isinstance(argv, (list, tuple)) or not argv:
            raise InvalidContentError(
                raw=False,
                reason="command content_ref requires non-empty argv list",
            )
        argv_tuple = tuple(str(item) for item in argv)
        env = value.get("env")
        if env is not None and not isinstance(env, Mapping):
            raise InvalidContentError(
                raw=False,
                reason="command content_ref env must be a mapping when provided",
            )
        cwd = value.get("cwd")
        if cwd is not None and not isinstance(cwd, str):
            raise InvalidContentError(
                raw=False,
                reason="command content_ref cwd must be a string when provided",
            )
        timeout = value.get("timeout_seconds", 30.0)
        if not isinstance(timeout, (int, float)) or isinstance(timeout, bool):
            raise InvalidContentError(
                raw=False,
                reason="command content_ref timeout_seconds must be a number",
            )
        return ContentSourceCommand(
            argv=argv_tuple,
            timeout_seconds=float(timeout),
            env=dict(env) if isinstance(env, Mapping) else None,
            cwd=cwd,
        )
    raise InvalidContentError(
        raw=False,
        reason=f"unknown content_ref type: {kind!r}",
    )


def _required_str(mapping: Mapping[str, Any], key: str) -> str:
    value = mapping.get(key)
    if not isinstance(value, str) or not value:
        raise InvalidContentError(
            raw=False,
            reason=f"content_ref missing required string field {key!r}",
        )
    return value


def _coerce_items(
    items: Iterable[str | ContentItem | Mapping[str, Any]],
) -> tuple[ContentItem, ...]:
    out = tuple(ContentItem.of(item) for item in items)
    if not out:
        raise InvalidContentError(raw=False, reason="items must be non-empty")
    return out


@dataclass(frozen=True, slots=True)
class InsertParaBefore:
    """Insert one or more paragraphs immediately before a target paragraph."""

    op_id: str
    target_id: str
    items: tuple[ContentItem, ...]
    raw: bool = False
    op: ClassVar[Literal["insert_para_before"]] = "insert_para_before"


@dataclass(frozen=True, slots=True)
class InsertParaAfter:
    """Insert one or more paragraphs immediately after a target paragraph."""

    op_id: str
    target_id: str
    items: tuple[ContentItem, ...]
    raw: bool = False
    op: ClassVar[Literal["insert_para_after"]] = "insert_para_after"


@dataclass(frozen=True, slots=True)
class ReplacePara:
    """Replace one paragraph with one or more new paragraphs."""

    op_id: str
    target_id: str
    items: tuple[ContentItem, ...]
    raw: bool = False
    op: ClassVar[Literal["replace_para"]] = "replace_para"


@dataclass(frozen=True, slots=True)
class DeletePara:
    """Delete one or more paragraphs referenced by ID."""

    op_id: str
    target_ids: tuple[str, ...]
    op: ClassVar[Literal["delete_para"]] = "delete_para"

    def __post_init__(self) -> None:
        if not self.target_ids:
            raise InvalidContentError(raw=False, reason="delete_para requires target_ids")
        if len(set(self.target_ids)) != len(self.target_ids):
            raise InvalidContentError(raw=False, reason="delete_para target_ids must be unique")


@dataclass(frozen=True, slots=True)
class ReplaceText:
    """Replace a text selector match inside a paragraph."""

    op_id: str
    target_id: str
    find: Selector
    content_literal: str | None = None
    content_ref: ContentRef | None = None
    occurrence: int | None = None
    op: ClassVar[Literal["replace_text"]] = "replace_text"

    def __post_init__(self) -> None:
        _require_single_content(
            raw=False,
            content_literal=self.content_literal,
            content_ref=self.content_ref,
            op=self.op,
        )


@dataclass(frozen=True, slots=True)
class DeleteText:
    """Delete a text selector match inside a paragraph."""

    op_id: str
    target_id: str
    find: Selector
    occurrence: int | None = None
    op: ClassVar[Literal["delete_text"]] = "delete_text"


@dataclass(frozen=True, slots=True)
class InsertTextBefore:
    """Insert text immediately before a selector match inside a paragraph."""

    op_id: str
    target_id: str
    find: Selector
    content_literal: str | None = None
    content_ref: ContentRef | None = None
    occurrence: int | None = None
    op: ClassVar[Literal["insert_text_before"]] = "insert_text_before"

    def __post_init__(self) -> None:
        _require_single_content(
            raw=False,
            content_literal=self.content_literal,
            content_ref=self.content_ref,
            op=self.op,
        )


@dataclass(frozen=True, slots=True)
class InsertTextAfter:
    """Insert text immediately after a selector match inside a paragraph."""

    op_id: str
    target_id: str
    find: Selector
    content_literal: str | None = None
    content_ref: ContentRef | None = None
    occurrence: int | None = None
    op: ClassVar[Literal["insert_text_after"]] = "insert_text_after"

    def __post_init__(self) -> None:
        _require_single_content(
            raw=False,
            content_literal=self.content_literal,
            content_ref=self.content_ref,
            op=self.op,
        )


AnyEditOperation: TypeAlias = (
    InsertParaBefore
    | InsertParaAfter
    | ReplacePara
    | DeletePara
    | ReplaceText
    | DeleteText
    | InsertTextBefore
    | InsertTextAfter
)


def _require_single_content(
    *,
    raw: bool,
    content_literal: str | None,
    content_ref: ContentRef | None,
    op: str,
) -> None:
    if (content_literal is None) == (content_ref is None):
        raise InvalidContentError(
            raw=raw,
            reason=f"{op} requires exactly one of content_literal or content_ref",
        )


class EditOperation:
    """Factory facade that produces concrete edit-operation dataclasses."""

    def __init__(self) -> None:  # pragma: no cover - defensive
        raise TypeError("EditOperation is a factory facade; use its classmethods")

    @staticmethod
    def insert_para_before(
        *,
        op_id: str,
        target_id: str,
        items: Iterable[str | ContentItem | Mapping[str, Any]],
        raw: bool = False,
    ) -> InsertParaBefore:
        return InsertParaBefore(
            op_id=op_id,
            target_id=target_id,
            items=_coerce_items(items),
            raw=raw,
        )

    @staticmethod
    def insert_para_after(
        *,
        op_id: str,
        target_id: str,
        items: Iterable[str | ContentItem | Mapping[str, Any]],
        raw: bool = False,
    ) -> InsertParaAfter:
        return InsertParaAfter(
            op_id=op_id,
            target_id=target_id,
            items=_coerce_items(items),
            raw=raw,
        )

    @staticmethod
    def replace_para(
        *,
        op_id: str,
        target_id: str,
        items: Iterable[str | ContentItem | Mapping[str, Any]],
        raw: bool = False,
    ) -> ReplacePara:
        return ReplacePara(
            op_id=op_id,
            target_id=target_id,
            items=_coerce_items(items),
            raw=raw,
        )

    @staticmethod
    def delete_para(*, op_id: str, target_ids: Iterable[str]) -> DeletePara:
        return DeletePara(op_id=op_id, target_ids=tuple(target_ids))

    @staticmethod
    def replace_text(
        *,
        op_id: str,
        paragraph_id: str,
        find: str | Selector | Mapping[str, Any],
        replacement: str | ContentItem | Mapping[str, Any],
        occurrence: int | None = None,
    ) -> ReplaceText:
        item = ContentItem.of(replacement)
        return ReplaceText(
            op_id=op_id,
            target_id=paragraph_id,
            find=Selector.coerce(find),
            content_literal=item.content_literal,
            content_ref=item.content_ref,
            occurrence=occurrence,
        )

    @staticmethod
    def delete_text(
        *,
        op_id: str,
        paragraph_id: str,
        find: str | Selector | Mapping[str, Any],
        occurrence: int | None = None,
    ) -> DeleteText:
        return DeleteText(
            op_id=op_id,
            target_id=paragraph_id,
            find=Selector.coerce(find),
            occurrence=occurrence,
        )

    @staticmethod
    def insert_text_before(
        *,
        op_id: str,
        paragraph_id: str,
        find: str | Selector | Mapping[str, Any],
        text: str | ContentItem | Mapping[str, Any],
        occurrence: int | None = None,
    ) -> InsertTextBefore:
        item = ContentItem.of(text)
        return InsertTextBefore(
            op_id=op_id,
            target_id=paragraph_id,
            find=Selector.coerce(find),
            content_literal=item.content_literal,
            content_ref=item.content_ref,
            occurrence=occurrence,
        )

    @staticmethod
    def insert_text_after(
        *,
        op_id: str,
        paragraph_id: str,
        find: str | Selector | Mapping[str, Any],
        text: str | ContentItem | Mapping[str, Any],
        occurrence: int | None = None,
    ) -> InsertTextAfter:
        item = ContentItem.of(text)
        return InsertTextAfter(
            op_id=op_id,
            target_id=paragraph_id,
            find=Selector.coerce(find),
            content_literal=item.content_literal,
            content_ref=item.content_ref,
            occurrence=occurrence,
        )


@dataclass(frozen=True, slots=True)
class OperationResult:
    """Successful outcome of one operation, with warnings and previews."""

    op_id: str
    op: str
    status: Literal["success"] = "success"
    target_id: str | None = None
    target_ids: tuple[str, ...] = ()
    new_ids: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()
    before_preview: str | None = None
    after_previews: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class EditResult:
    """Successful batch result. Failures are raised, never returned here."""

    results: tuple[OperationResult, ...] = field(default_factory=tuple)


@dataclass(frozen=True, slots=True)
class SaveResult:
    """Result of ``Document.save``: output path, optional backup, warnings."""

    output_path: str
    backup_path: str | None
    warnings: tuple[str, ...] = ()


__all__ = [
    "AnyEditOperation",
    "ContentItem",
    "ContentRef",
    "ContentSourceCommand",
    "ContentSourceFile",
    "ContentSourceJsonPath",
    "DeletePara",
    "DeleteText",
    "EditOperation",
    "EditResult",
    "InsertParaAfter",
    "InsertParaBefore",
    "InsertTextAfter",
    "InsertTextBefore",
    "OperationResult",
    "Pagination",
    "ParagraphId",
    "ParagraphInfo",
    "ParagraphListResult",
    "ParagraphLocation",
    "ParagraphMatchInfo",
    "ParagraphSearchResult",
    "ReplacePara",
    "ReplaceText",
    "SaveResult",
    "Selector",
    "TableContext",
    "TextMatch",
]
