# Errors

Every failure is a subclass of
[`DocxKnifeError`](api.md#docx_knife.DocxKnifeError) and carries
structured, JSON-serializable fields exposed via `to_dict()`. Message
strings are always bounded, so error output cannot blow up logs or LLM
contexts.

## Reference table

| Error | Fields | When it fires |
| --- | --- | --- |
| [`DocxKnifeError`](api.md#docx_knife.DocxKnifeError) | (base) | Base class for every failure. Catch it to handle any package error. |
| [`DocumentNotFoundError`](api.md#docx_knife.DocumentNotFoundError) | `path` | `Document.open` on a missing path. |
| [`InvalidDocumentError`](api.md#docx_knife.InvalidDocumentError) | `path`, `reason` | `Document.open` on a bad ZIP, a package missing `word/document.xml`, or an unparseable main XML. |
| [`SourceChangedError`](api.md#docx_knife.SourceChangedError) | `source_path` | `Document.save` when the source file bytes / size / mtime differ from the fingerprint captured at `open`. |
| [`ParagraphNotFoundError`](api.md#docx_knife.ParagraphNotFoundError) | `target_id` | Any API given an unknown ID, or an ID that a prior operation invalidated (e.g. via `delete_para` / `replace_para`). |
| [`TextNotFoundError`](api.md#docx_knife.TextNotFoundError) | `target_id`, `selector`, `occurrence`, `total_matches` | A selector matched zero times, or `occurrence` was out of range for the number of matches. |
| [`AmbiguousTextMatchError`](api.md#docx_knife.AmbiguousTextMatchError) | `target_id`, `selector`, `total_matches` | A selector matched more than once and `occurrence` was omitted. |
| [`InvalidPatternError`](api.md#docx_knife.InvalidPatternError) | `pattern`, `reason` | An empty literal pattern, or a regex that failed to compile. |
| [`InvalidContentError`](api.md#docx_knife.InvalidContentError) | `raw`, `reason` | A `ContentItem` violates source cardinality, a `content_ref` points outside the input roots, decoding fails, a command times out or exits non-zero, jsonpath returned zero or multiple values, or a raw fragment is malformed. |
| [`UnsupportedStructureError`](api.md#docx_knife.UnsupportedStructureError) | `target_id`, `structures`, `matched_range` | A text edit range crosses a protected structure (`w:fldChar`, `w:hyperlink`, `w:bookmarkStart/End`, `w:ins`, `w:del`, ...) or hits an atomic reserved-marker range (`w:tab`, `w:br`, `w:cr`) that the executor refuses to break. |
| [`BatchOperationError`](api.md#docx_knife.BatchOperationError) | `operation_index`, `op_id`, `reason`, `rolled_back` (always `True`) | Any operation inside `batch_edit` fails. The document is guaranteed to be exactly as it was before the batch. `__cause__` carries the original underlying error. |
| [`ValidationError`](api.md#docx_knife.ValidationError) | `stage`, `checks`, `failed_check` | Schema validation (`stage="schema"`), batch prevalidation (`stage="prevalidation"`), precommit invariant checks (`stage="precommit"`), or save-time package checks (`stage="save"`). |

## Handling patterns

### Rebuild after invalidation

`ParagraphNotFoundError` is the correct signal that IDs from a prior
session are stale, or that the batch you assembled overlaps a
prior invalidation. Refetch IDs before rebuilding:

```python
from docx_knife import Document, ParagraphNotFoundError

with Document.open("in.docx") as doc:
    try:
        doc.batch_edit(operations)
    except ParagraphNotFoundError as err:
        stale_id = err.target_id
        # Re-query and rebuild — never invent replacement IDs.
        ...
```

### Selector ambiguity

```python
from docx_knife import AmbiguousTextMatchError, EditOperation

try:
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op",
            paragraph_id=pid,
            find="target",
            replacement="new",
        ),
    ])
except AmbiguousTextMatchError as err:
    print(err.total_matches, "matches; supply occurrence or narrow selector")
```

### Atomic rollback

```python
from docx_knife import BatchOperationError

try:
    doc.batch_edit(operations)
except BatchOperationError as err:
    assert err.rolled_back is True
    # Inspect err.op_id / err.reason / err.__cause__ to fix the batch.
    ...
```

## Serialization

Every error exposes `to_dict()`, which returns a plain dict suitable for
JSON logging:

```python
try:
    doc.batch_edit(operations)
except DocxKnifeError as err:
    log.error("docx-knife failure", extra=err.to_dict())
```

Selector and tuple fields are recursively encoded, and all string
previews are truncated to 80 characters.
