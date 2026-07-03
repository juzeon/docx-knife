# Errors

Every failure subclasses [`DocxKnifeError`](api.md#docx_knife.DocxKnifeError) and carries structured JSON-serializable fields via `to_dict()`. Message strings are bounded so error output cannot blow up logs or LLM contexts.

## Reference table

| Error | Fields | Fires when |
| --- | --- | --- |
| `DocumentNotFoundError` | `path` | `Document.open` on missing path. |
| `InvalidDocumentError` | `path`, `reason` | Bad ZIP, missing `word/document.xml`, or unparseable main XML. |
| `SourceChangedError` | `source_path` | `save` sees source bytes/size/mtime differ from the `open` fingerprint. |
| `ParagraphNotFoundError` | `target_id` | Unknown ID, or ID invalidated by a prior `delete_para` / `replace_para`. |
| `TextNotFoundError` | `target_id`, `selector`, `occurrence`, `total_matches` | Selector matched zero times or `occurrence` out of range. |
| `AmbiguousTextMatchError` | `target_id`, `selector`, `total_matches` | Selector matched > 1 and `occurrence` was omitted. |
| `InvalidPatternError` | `pattern`, `reason` | Empty literal pattern, or regex failed to compile. |
| `InvalidContentError` | `raw`, `reason` | Cardinality violation, path outside input roots, decode failure, command timeout/non-zero exit, jsonpath 0-or-many hits, malformed raw fragment. |
| `UnsupportedStructureError` | `target_id`, `structures`, `matched_range` | Text edit crosses a protected structure (`w:fldChar`, `w:hyperlink`, `w:bookmarkStart/End`, `w:ins`, `w:del`, ...) or an atomic reserved-marker range (`w:tab`, `w:br`, `w:cr`). |
| `BatchOperationError` | `operation_index`, `op_id`, `reason`, `rolled_back` (always `True`) | Any op inside `batch_edit` fails. Document state is exactly the pre-batch state. `__cause__` carries the underlying error. |
| `ValidationError` | `stage`, `checks`, `failed_check` | `stage="schema"` (JSON schema), `"prevalidation"`, `"precommit"`, or `"save"`. |

## Handling patterns

### Rebuild after invalidation

```python
from docx_knife import Document, ParagraphNotFoundError

with Document.open("in.docx") as doc:
    try:
        doc.batch_edit(operations)
    except ParagraphNotFoundError as err:
        stale_id = err.target_id
        # Re-query and rebuild — never invent replacement IDs.
```

### Selector ambiguity

```python
from docx_knife import AmbiguousTextMatchError, EditOperation

try:
    doc.batch_edit([EditOperation.replace_text(
        op_id="op", paragraph_id=pid, find="target", replacement="new",
    )])
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
```

## Serialization

```python
try:
    doc.batch_edit(operations)
except DocxKnifeError as err:
    log.error("docx-knife failure", extra=err.to_dict())
```

Selector and tuple fields are recursively encoded; all string previews are truncated to 80 characters.
