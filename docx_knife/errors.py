"""Public exception hierarchy for ``docx_knife``.

Every exception inherits :class:`DocxKnifeError` and carries structured,
JSON-serializable fields under :meth:`to_dict`. Human-facing message text is
always bounded so that stray previews cannot blow up logs or LLM contexts.

Design notes:

* Exceptions are hand-written (no ``@dataclass`` decorator) because the
  dataclass machinery interferes with ``BaseException``'s pickling contract.
* Public fields are declared via the ``__public_fields__`` class attribute,
  which the ``_Serializable`` mixin uses to build ``to_dict``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, ClassVar

if TYPE_CHECKING:
    from ._models import Selector

_PREVIEW_LIMIT: int = 80


def _truncate_preview(value: str, limit: int = _PREVIEW_LIMIT) -> str:
    """Bound a human-readable preview so error messages stay short."""
    if len(value) <= limit:
        return value
    return value[: max(0, limit - 1)] + "\u2026"


def _selector_payload(selector: Selector) -> dict[str, Any]:
    return {"pattern": _truncate_preview(selector.pattern), "regex": selector.regex}


class _Serializable:
    """Mixin that exposes public dataclass-like fields as ``to_dict``."""

    __public_fields__: ClassVar[tuple[str, ...]] = ()

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"code": getattr(self, "code", self.__class__.__name__)}
        for field in self.__public_fields__:
            payload[field] = _encode(getattr(self, field))
        return payload


def _encode(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    from ._models import Selector as _Selector

    if isinstance(value, _Selector):
        return _selector_payload(value)
    if isinstance(value, (list, tuple)):
        return [_encode(item) for item in value]
    if isinstance(value, dict):
        return {str(k): _encode(v) for k, v in value.items()}
    return str(value)


class DocxKnifeError(_Serializable, Exception):
    """Base class for every error raised by ``docx_knife``."""

    code: ClassVar[str] = "docx_knife_error"


class DocumentNotFoundError(DocxKnifeError):
    """Raised by :meth:`Document.open` when the source path is missing."""

    code: ClassVar[str] = "document_not_found"
    __public_fields__: ClassVar[tuple[str, ...]] = ("path",)

    def __init__(self, *, path: str) -> None:
        self.path = path
        super().__init__(f"docx source not found: {path}")


class InvalidDocumentError(DocxKnifeError):
    """Raised when the source is not a valid DOCX (bad ZIP or unparseable XML)."""

    code: ClassVar[str] = "invalid_document"
    __public_fields__: ClassVar[tuple[str, ...]] = ("path", "reason")

    def __init__(self, *, path: str, reason: str) -> None:
        self.path = path
        self.reason = reason
        super().__init__(f"invalid docx {path}: {_truncate_preview(reason)}")


class SourceChangedError(DocxKnifeError):
    """Raised on save when the source file has been modified on disk since open."""

    code: ClassVar[str] = "source_changed"
    __public_fields__: ClassVar[tuple[str, ...]] = ("source_path",)

    def __init__(self, *, source_path: str) -> None:
        self.source_path = source_path
        super().__init__(f"source file changed on disk since open: {source_path}")


class ParagraphNotFoundError(DocxKnifeError):
    """Raised when a paragraph ID is unknown or has been invalidated."""

    code: ClassVar[str] = "paragraph_not_found"
    __public_fields__: ClassVar[tuple[str, ...]] = ("target_id",)

    def __init__(self, *, target_id: str) -> None:
        self.target_id = target_id
        super().__init__(f"paragraph id not found: {target_id}")


class TextNotFoundError(DocxKnifeError):
    """Raised when a selector has no match, or the requested occurrence is out of range."""

    code: ClassVar[str] = "text_not_found"
    __public_fields__: ClassVar[tuple[str, ...]] = (
        "target_id",
        "selector",
        "occurrence",
        "total_matches",
    )

    def __init__(
        self,
        *,
        target_id: str,
        selector: Selector,
        occurrence: int | None,
        total_matches: int,
    ) -> None:
        self.target_id = target_id
        self.selector = selector
        self.occurrence = occurrence
        self.total_matches = total_matches
        super().__init__(
            f"selector {_truncate_preview(selector.pattern)!r} did not match in {target_id} "
            f"(occurrence={occurrence}, total_matches={total_matches})"
        )


class AmbiguousTextMatchError(DocxKnifeError):
    """Raised when a selector without an ``occurrence`` matches more than once."""

    code: ClassVar[str] = "ambiguous_text_match"
    __public_fields__: ClassVar[tuple[str, ...]] = ("target_id", "selector", "total_matches")

    def __init__(
        self,
        *,
        target_id: str,
        selector: Selector,
        total_matches: int,
    ) -> None:
        self.target_id = target_id
        self.selector = selector
        self.total_matches = total_matches
        super().__init__(
            f"selector {_truncate_preview(selector.pattern)!r} matched {total_matches} times "
            f"in {target_id}; specify occurrence"
        )


class InvalidPatternError(DocxKnifeError):
    """Raised for an unusable literal (empty) or regex (compile error) selector."""

    code: ClassVar[str] = "invalid_pattern"
    __public_fields__: ClassVar[tuple[str, ...]] = ("pattern", "reason")

    def __init__(self, *, pattern: str, reason: str) -> None:
        self.pattern = pattern
        self.reason = reason
        super().__init__(
            f"invalid selector pattern {_truncate_preview(pattern)!r}: {_truncate_preview(reason)}"
        )


class InvalidContentError(DocxKnifeError):
    """Raised when supplied content or a content reference is malformed."""

    code: ClassVar[str] = "invalid_content"
    __public_fields__: ClassVar[tuple[str, ...]] = ("raw", "reason")

    def __init__(self, *, raw: bool, reason: str) -> None:
        self.raw = raw
        self.reason = reason
        mode = "raw" if raw else "visible"
        super().__init__(f"invalid {mode}-mode content: {_truncate_preview(reason)}")


class UnsupportedStructureError(DocxKnifeError):
    """Raised when a text match crosses protected or atomic structures."""

    code: ClassVar[str] = "unsupported_structure"
    __public_fields__: ClassVar[tuple[str, ...]] = (
        "target_id",
        "structures",
        "matched_range",
    )

    def __init__(
        self,
        *,
        target_id: str,
        structures: tuple[str, ...],
        matched_range: tuple[int, int] | None,
    ) -> None:
        self.target_id = target_id
        self.structures = structures
        self.matched_range = matched_range
        super().__init__(
            f"selector in {target_id} crosses unsupported structures: {list(structures)}"
        )


class BatchOperationError(DocxKnifeError):
    """Raised when a batch fails; the document has been rolled back."""

    code: ClassVar[str] = "batch_operation_error"
    __public_fields__: ClassVar[tuple[str, ...]] = (
        "operation_index",
        "op_id",
        "reason",
        "rolled_back",
    )

    def __init__(
        self,
        *,
        operation_index: int,
        op_id: str,
        reason: str,
        cause: BaseException | None = None,
        rolled_back: bool = True,
    ) -> None:
        self.operation_index = operation_index
        self.op_id = op_id
        self.reason = reason
        self.rolled_back = rolled_back
        super().__init__(
            f"batch failed at index {operation_index} ({op_id}): {_truncate_preview(reason)}"
        )
        if cause is not None:
            self.__cause__ = cause

    def __reduce__(self) -> tuple[Any, ...]:
        return (
            _rebuild_batch_error,
            (self.operation_index, self.op_id, self.reason, self.rolled_back),
        )


def _rebuild_batch_error(
    operation_index: int, op_id: str, reason: str, rolled_back: bool
) -> BatchOperationError:
    return BatchOperationError(
        operation_index=operation_index,
        op_id=op_id,
        reason=reason,
        rolled_back=rolled_back,
    )


class ValidationError(DocxKnifeError):
    """Raised by schema, prevalidation, and precommit checks."""

    code: ClassVar[str] = "validation_error"
    __public_fields__: ClassVar[tuple[str, ...]] = ("stage", "checks", "failed_check")

    def __init__(self, *, stage: str, checks: tuple[str, ...], failed_check: str) -> None:
        self.stage = stage
        self.checks = checks
        self.failed_check = failed_check
        super().__init__(f"validation failed at stage {stage!r}: {failed_check}")


__all__ = [
    "AmbiguousTextMatchError",
    "BatchOperationError",
    "DocumentNotFoundError",
    "DocxKnifeError",
    "InvalidContentError",
    "InvalidDocumentError",
    "InvalidPatternError",
    "ParagraphNotFoundError",
    "SourceChangedError",
    "TextNotFoundError",
    "UnsupportedStructureError",
    "ValidationError",
]
