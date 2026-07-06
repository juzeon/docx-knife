# docx-knife

AI-safe DOCX patch engine with stable paragraph anchors.

**docx-knife** lets AI agents (and humans) edit `.docx` files without hallucinating text, breaking OOXML structure, or loading the whole document into context. It assigns stable IDs to every paragraph, resolves text across fragmented runs, and applies edits as atomic batches with full rollback on failure.

## Why?

When LLMs edit Word documents they repeatedly:

- Hallucinate text when asked to reproduce large passages
- Break OOXML structure (Word shows "needs repair")
- Corrupt formatting by merging runs
- Miss targets because matching relies on exact text reproduction

docx-knife solves these by **decoupling location from content**: paragraphs are addressed by stable IDs, text is located via TextMap across run boundaries, and edits are structured operations—not free-form rewrites.

## Installation

```bash
pip install docx-knife
```

Requires Python 3.10+.

## Quick Start

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    # Browse paragraphs
    listing = doc.list_paragraphs(start=1, limit=20)
    for para in listing.paragraphs:
        print(f"{para.id}: {para.text[:60]}")

    # Find and replace text (works even across fragmented runs)
    target = listing.paragraphs[0]
    doc.batch_edit(operations=[
        EditOperation.replace_text(
            paragraph_id=target.id,
            find="thirty days",
            replacement="sixty days",
            occurrence=0,
        )
    ])

    doc.save("edited.docx")
```

## Key Features

### Stable Paragraph Anchors

Every `<w:p>` gets an engine-assigned ID (`p_000001`, `p_000002`, …) on parse. IDs survive edits to other paragraphs, and new paragraphs receive monotonically increasing IDs. No reliance on `w14:paraId` or XML indices.

### Cross-Run Text Matching

TextMap concatenates visible text from all `<w:t>` nodes in a paragraph and maps every character back to its source node. "违约责任" split across two runs? Still a single match target.

### Atomic Batch Operations

All edits in a batch either succeed together or roll back completely—DOM, anchor manifest, and change log all revert. No half-applied states.

### Paragraph Operations

```python
doc.insert_para_after(anchor, ["New paragraph A", "New paragraph B"])
doc.insert_para_before(anchor, ["Before content"])
doc.replace_para(target.id, ["Replacement line 1", "Replacement line 2"])
doc.delete_para([obsolete.id])
```

### In-Paragraph Text Operations

```python
doc.batch_edit(operations=[
    EditOperation.replace_text(paragraph_id=p.id, find="old", replacement="new"),
    EditOperation.delete_text(paragraph_id=p.id, find="remove this"),
    EditOperation.insert_text_after(paragraph_id=p.id, find="anchor", text=" appended"),
    EditOperation.insert_text_before(paragraph_id=p.id, find="anchor", text="prepended "),
])
```

### Search and Discovery

```python
doc.paragraph_count()
doc.grep_paragraphs("penalty clause", regex=False)
doc.find_text("违约金", paragraph_id=target.id)
doc.count_matches(r"第[一二三四五]条", regex=True, paragraph_id=target.id)
```

### Safe Save with Backup

```python
result = doc.save("output.docx")
print(result.backup_path)  # output.docx.bak (if overwriting)
```

Detects external modifications to the source file and refuses to silently overwrite.

### Raw XML Mode

For advanced use cases, read and write raw `<w:p>` XML without text projection:

```python
xml = doc.get_paragraph(target.id, raw=True)
doc.replace_para(target.id, [modified_xml], raw=True)
```

## Table Support

Paragraphs inside table cells are fully supported with the same ID and operation API. Location metadata includes table index, row, column (accounting for grid spans), and nesting depth.

## Error Handling

All exceptions inherit from `DocxKnifeError` with structured, serializable fields:

| Exception | When |
|---|---|
| `DocumentNotFoundError` | Source path doesn't exist |
| `InvalidDocumentError` | Not a valid DOCX or corrupt XML |
| `ParagraphNotFoundError` | Target ID not in manifest |
| `TextNotFoundError` | Pattern not found in target paragraph |
| `AmbiguousTextMatchError` | Multiple matches without explicit `occurrence` |
| `BatchOperationError` | Any operation in a batch fails (auto-rollback) |
| `SourceChangedError` | Source file modified externally before save |
| `ValidationError` | Post-edit XML validation failure |

## JSON Batch Format

For agent integration, edits can be expressed as JSON:

```json
{
  "operations": [
    {
      "op_id": "op_001",
      "op": "replace_text",
      "target_id": "p_000042",
      "find": "三十日",
      "occurrence": 0,
      "content_literal": "六十日"
    }
  ]
}
```

## License

MIT
