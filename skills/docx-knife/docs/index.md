# docx-knife

AI-safe DOCX patch engine. The LLM never emits XML, XPath, indexes, or `w14:paraId` — only paragraph IDs returned by read APIs plus structured operations against them. The engine owns OOXML parsing, cross-run text mapping, structure protection, atomic rollback, and safe save.

## Minimal example

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    target = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op_001",
            paragraph_id=target,
            find="三十日",
            replacement="六十日",
        ),
    ])
    saved = doc.save("contract.edited.docx")
    print(saved.backup_path)  # -> contract.edited.docx.bak, or None
```

Next: [Quickstart](quickstart.md) · [API reference](api.md) · [Errors](errors.md).
