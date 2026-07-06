---
name: docx-knife
description: Edit .docx files without hallucinating text, breaking OOXML, or dumping the whole document into context. Use for any request that mutates paragraphs or paragraph-internal spans in a DOCX file.
---

# docx-knife

Query paragraph IDs, then submit one atomic batch referencing those IDs. The engine owns XML parsing, TextMap alignment, structural preservation, rollback, and safe save.

## Rules

- Only use paragraph IDs returned by read APIs **in this session**. Never invent IDs.
- Strings > ~100 chars → `content_ref`. Short human phrases ≤ ~40 chars → `content_literal`.
- `raw=True` is allowed on paragraph-level ops (`insert_para_before`, `insert_para_after`, `replace_para`) when you need to emit exact OOXML `<w:p>` fragments; paragraph-internal text ops reject it.
- One `batch_edit` = one atomic write. Never call `save()` inside a batch.
- On failure: read the error, fix the batch, resubmit. Never resubmit verbatim.
- Keep batches ≤ 50 operations.

## Examples

### Find and replace text in a paragraph

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    target = doc.grep_paragraphs("30 天").matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op1", paragraph_id=target,
            find="30 天", replacement="60 天",
        ),
    ])
    doc.save("contract.edited.docx")
```

### List and browse paragraphs

```python
with Document.open("report.docx") as doc:
    # First 10 paragraphs, 80-char preview
    result = doc.list_paragraphs(start=1, limit=10, max_chars=80)
    for info in result.paragraphs:
        print(info.id, "|", info.text)

    # Search by pattern
    hits = doc.grep_paragraphs("违约", regex=False, max_chars=60)
    for m in hits.matches:
        print(m.paragraph.id, m.match_count, "|", m.paragraph.text)
```

### Multi-operation atomic batch

```python
with Document.open("contract.docx") as doc:
    anchor = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    target = doc.grep_paragraphs("责任条款 1").matches[0].paragraph.id
    stale = doc.list_paragraphs(start=doc.paragraph_count(), limit=1).paragraphs[0].id

    result = doc.batch_edit([
        EditOperation.insert_para_after(
            op_id="op1", target_id=anchor,
            items=["This clause was inserted by docx-knife."],
        ),
        EditOperation.replace_para(
            op_id="op2", target_id=target,
            items=["第 1 条：责任条款已被整段替换。"],
        ),
        EditOperation.delete_para(op_id="op3", target_ids=[stale]),
        EditOperation.replace_text(
            op_id="op4", paragraph_id=anchor,
            find="Test Contract", replacement="Sample Contract",
        ),
    ])
    doc.save("contract.edited.docx")
```

### Insert text before/after a match

```python
with Document.open("memo.docx") as doc:
    pid = doc.grep_paragraphs("deadline").matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.insert_text_after(
            op_id="op1", paragraph_id=pid,
            find="deadline", text=" (extended)",
        ),
    ])
    doc.save("memo.edited.docx")
```

### Delete text from a paragraph

```python
with Document.open("draft.docx") as doc:
    pid = doc.grep_paragraphs("DRAFT").matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.delete_text(op_id="op1", paragraph_id=pid, find="DRAFT"),
    ])
    doc.save("draft.clean.docx")
```

### Use `content_ref` for long or deterministic text

```python
from docx_knife import Document, EditOperation, ContentItem, ContentSourceJsonPath, ContentSourceFile

with Document.open("contract.docx") as doc:
    pid = doc.grep_paragraphs("Party A").matches[0].paragraph.id

    # From a JSON file
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op1", paragraph_id=pid,
            find="Party A",
            replacement=ContentItem(
                content_ref=ContentSourceJsonPath(source="data.json", path="$.party_a.name"),
            ),
        ),
    ])

    # Replace a whole paragraph with content from a file
    target = doc.grep_paragraphs("保密条款").matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.replace_para(
            op_id="op2", target_id=target,
            items=[ContentItem(content_ref=ContentSourceFile(path="clauses/confidentiality.txt"))],
        ),
    ])
    doc.save("contract.final.docx")
```

### Emit raw OOXML with `raw=True`

Use `raw=True` on paragraph-level ops when you need exact WordprocessingML control (custom run properties, fields, structured tags). Each item must be a top-level `<w:p>` in the standard `w:` namespace. Text expansion and normalization are bypassed.

```python
with Document.open("contract.docx") as doc:
    anchor = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    fragment = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">附录 A</w:t></w:r>'
        '</w:p>'
    )
    doc.batch_edit([
        EditOperation.insert_para_after(
            op_id="op1", target_id=anchor,
            items=[fragment], raw=True,
        ),
    ])
    doc.save("contract.edited.docx")
```

Only `insert_para_before`, `insert_para_after`, and `replace_para` accept `raw=True`. Paragraph-internal text ops (`replace_text`, `insert_text_*`, `delete_text`) reject it.

### Handle multiple matches with `occurrence`

```python
with Document.open("report.docx") as doc:
    pid = doc.grep_paragraphs("2024").matches[0].paragraph.id

    # Replace only the first occurrence (0-indexed)
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op1", paragraph_id=pid,
            find="2024", replacement="2025", occurrence=0,
        ),
    ])

    # Replace ALL occurrences (right-to-left, offsets stay valid)
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op2", paragraph_id=pid,
            find="2024", replacement="2025", occurrence=-1,
        ),
    ])
    doc.save("report.updated.docx")
```

### Regex selector

```python
with Document.open("legal.docx") as doc:
    pid = doc.grep_paragraphs("第.*条", regex=True).matches[0].paragraph.id
    doc.batch_edit([
        EditOperation.replace_text(
            op_id="op1", paragraph_id=pid,
            find={"pattern": r"第\d+条", "regex": True},
            replacement="第 99 条", occurrence=0,
        ),
    ])
    doc.save("legal.edited.docx")
```

### Error handling

```python
from docx_knife import (
    Document, EditOperation, BatchOperationError,
    AmbiguousTextMatchError, ParagraphNotFoundError,
)

with Document.open("doc.docx") as doc:
    try:
        doc.batch_edit(operations)
    except AmbiguousTextMatchError as e:
        # Multiple matches — supply occurrence or narrow the selector
        print(f"{e.total_matches} matches for selector on {e.target_id}")
    except ParagraphNotFoundError as e:
        # ID was invalidated — re-query
        print(f"stale ID: {e.target_id}")
    except BatchOperationError as e:
        # Entire batch rolled back; inspect and fix
        print(f"op {e.op_id} failed: {e.reason}")
        assert e.rolled_back is True
```

## Operation reference

| Op | Target | What it does |
| --- | --- | --- |
| `insert_para_before` | `target_id` | Insert paragraphs before anchor |
| `insert_para_after` | `target_id` | Insert paragraphs after anchor |
| `replace_para` | `target_id` | Replace entire paragraph (old ID invalidated) |
| `delete_para` | `target_ids` | Delete paragraphs (IDs invalidated) |
| `replace_text` | `paragraph_id` + `find` | Replace matched text span |
| `insert_text_before` | `paragraph_id` + `find` | Insert before matched span |
| `insert_text_after` | `paragraph_id` + `find` | Insert after matched span |
| `delete_text` | `paragraph_id` + `find` | Delete matched span |

## Key behaviors

- **Rollback**: any op failure rolls back the entire batch to pre-batch state.
- **ID invalidation**: `replace_para` and `delete_para` permanently invalidate target IDs. New IDs are returned in `OperationResult.new_ids`.
- **Style inheritance**: new paragraphs inherit `w:pPr` and the first ordinary text run's `w:rPr` from an anchor paragraph. `insert_para_after` / `replace_para` use the target as anchor; `insert_para_before` uses the *previous sibling paragraph* when one exists so inserting before a heading continues the body flow (falls back to the target when there is no previous paragraph). Text edits inherit `w:rPr` from the left boundary. Runs nested inside revision wrappers like `<w:ins>` count as ordinary text runs for this lookup.
- **Newlines in content**: `\n` → `<w:br/>`; `\n\n`+ → paragraph split.
- **`normalize_text=True`**: opt-in CJK punctuation normalization (half→full-width when adjacent to CJK).
- **Conflict rules**: no `replace_para` + text op on same target; no duplicate `op_id`; no op referencing an already-invalidated ID.

## Further reference

- Machine-readable schema: [agent_schema.json](agent_schema.json)
- Content sources detail: [docs/content-sources.md](docs/content-sources.md)
- Error fields and recovery patterns: [docs/errors.md](docs/errors.md)
- Operation conflict matrix: [docs/operations.md](docs/operations.md)
- Lifecycle invariants: [docs/lifecycle.md](docs/lifecycle.md)
