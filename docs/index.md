# docx-knife

`docx-knife` is an AI-safe DOCX patch engine. It solves a specific problem:
letting an LLM edit Word documents *without* asking the model to reproduce
long stretches of original text. The engine assigns every editable paragraph
a stable, instance-local ID, and the LLM only emits structured operations
that reference those IDs. The engine — not the model — owns OOXML parsing,
cross-run text mapping, structure protection, atomic rollback, and safe save.

## Why not just search-and-replace?

Straight search-replace on DOCX regularly breaks in ways that are hard to
diagnose:

| Symptom | Root cause |
| --- | --- |
| Long text hallucinated / reworded | The LLM was asked to reproduce a long span verbatim; autoregressive generation drifts. |
| Insert/replace at the wrong spot | The content string was used as the locator, so it had to be reproduced perfectly. |
| `old_str` fails as it grows | Locating and reproducing were coupled; the string had to be both unique and short. |
| CJK/ASCII punctuation drift | Punctuation width/type is deterministic from context but was left to the model. |
| Context window blown up | The whole document was pasted in instead of paged on demand. |
| "Word needs to repair this file" | The output was written back without OOXML validation; runs were merged unsafely. |

docx-knife decouples all of that. The LLM never sees XML, XPath, indexes, or
`w14:paraId`. It only sees paragraph IDs returned by the engine and the
`raw=false` operation schema.

## Minimal example

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    listing = doc.list_paragraphs(start=1, limit=50)
    target = listing.paragraphs[0]

    doc.batch_edit(
        operations=[
            EditOperation.replace_text(
                op_id="op_001",
                paragraph_id=target.id,
                find="三十日",
                replacement="六十日",
            )
        ],
    )
    saved = doc.save("contract.edited.docx")
    print(saved.backup_path)  # -> contract.edited.docx.bak, or None
```

See [Quickstart](quickstart.md) for an end-to-end example, and
[API Reference](api.md) for the complete public surface.
