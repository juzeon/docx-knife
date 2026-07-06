---
name: docx-knife
description: Edit .docx files without hallucinating text, breaking OOXML, or dumping the whole document into context. Use for any request that mutates paragraphs or paragraph-internal spans in a DOCX file.
---

# docx-knife

Query for paragraph IDs, then submit one atomic batch of operations that reference those IDs. The engine owns XML parsing, TextMap alignment, structural preservation, rollback, and safe save. Read further docs only when needed.

## Non-negotiable rules

- Only paragraph IDs returned by read APIs in **this session** are valid. Never invent IDs, XPath, indexes, `w14:paraId`, or offsets.
- Any string longer than ~100 chars must go through `content_ref`, not `content_literal`.
- `raw=true` is forbidden. The LLM-facing schema rejects the field.
- One `batch_edit` is one atomic write; do not call `save()` inside a batch, and never resubmit a failed batch verbatim — read the error, fix the batch, then resubmit.
- Keep a batch to ≤ 50 operations; split larger changes into successive atomic batches.

## Workflow

1. **Locate.** `list_paragraphs` / `grep_paragraphs` / `find_text` → paragraph IDs + previews.
2. **Compose.** Build operations referencing those IDs; prefer `content_ref` for deterministic or long text.
3. **Submit.** `batch_edit(operations=[...])`. Success → `EditResult`. Failure → structured `DocxKnifeError`.
4. **Save.** `save(output_path)` once. Writes atomically, produces `.bak` if the destination existed.

Reopening the file (new `Document.open`) invalidates all prior IDs.

## Public API (one-line signatures)

Read:

- `Document.open(source_path, *, content_config=None) -> Document`  (also `with ...:`)
- `Document.paragraph_count()`, `list_paragraphs`, `get_paragraph`, `get_visible_text`
- `Document.grep_paragraphs`, `count_matches`, `find_text`, `get_paragraph_object`

Write (one batch per call):

- `Document.batch_edit(operations, *, normalize_text=False, envelope=None) -> EditResult`
- `Document.save(output_path) -> SaveResult`
- `Document.change_log()` — audit only.

Schema helpers: `docx_knife.BATCH_SCHEMA`, `docx_knife.validate_batch(payload)`.

## Minimal example

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    target = doc.grep_paragraphs("30 天").matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.replace_text(op_id="op1", paragraph_id=target,
                                   find="30 天", replacement="60 天"),
    ])
    doc.save("contract.edited.docx")
```

End-to-end batch (replace + insert + delete in one call): see [docs/quickstart.md](docs/quickstart.md).

## Batch JSON schema

Every operation has a unique `op_id`, an `op` discriminator, target IDs, and either `items` (paragraph ops) or `find` + content (text ops). `additionalProperties=false` everywhere; `raw` is not permitted.

Machine-readable schema: [`agent_schema.json`](agent_schema.json). Regenerate via `python skills/docx-knife/_export_schema.py`.

Concrete envelopes: [`examples/replace_dates.json`](examples/replace_dates.json), [`examples/insert_and_delete.json`](examples/insert_and_delete.json).

## When to read what

- Op semantics, conflict matrix, `w:pPr` inheritance → [docs/operations.md](docs/operations.md)
- `content_literal` vs. `content_ref` (jsonpath / file / command), newline expansion, `normalize_text`, selectors, raw mode → [docs/content-sources.md](docs/content-sources.md)
- Structured error fields and recommended responses → [docs/errors.md](docs/errors.md)
- Open / query / batch / save invariants and fingerprinting → [docs/lifecycle.md](docs/lifecycle.md)
- Full API reference → [agent_schema.json](agent_schema.json)
