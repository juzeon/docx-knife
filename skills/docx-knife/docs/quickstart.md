# Quickstart

## 1. Install

```bash
pip install docx-knife
```

## 2. Open a document

`Document.open` extracts the package into a private per-instance workspace, captures a source fingerprint (used by `save` to detect drift), and returns a context manager.

```python
from docx_knife import Document

with Document.open("contract.docx") as doc:
    print(f"paragraphs: {doc.paragraph_count()}")
```

## 3. Discover paragraph IDs

Every batch operation must reference an ID returned by a read API in the current session.

```python
result = doc.list_paragraphs(start=1, limit=5, max_chars=40)
for info in result.paragraphs:
    print(info.id, "|", info.text)
```

```text
p_000001 | 契约合同 — Test Contract
p_000002 | 第一部分 定义
p_000003 | 第 1 条：定义条款 1 的正文内容。
```

Locate by pattern with `grep_paragraphs`:

```python
hits = doc.grep_paragraphs("target", regex=False, max_chars=60)
for match in hits.matches:
    print(match.paragraph.id, match.match_count, "|", match.paragraph.text)
```

## 4. One atomic batch

One `batch_edit` is one transaction: any failure rolls the DOM back exactly.

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    target = doc.grep_paragraphs("责任条款 1").matches[0].paragraph.id
    anchor = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    stale = doc.list_paragraphs(start=doc.paragraph_count(), limit=1).paragraphs[0].id

    result = doc.batch_edit([
        EditOperation.insert_para_after(
            op_id="op_insert", target_id=anchor,
            items=["This clause was inserted by docx-knife."],
        ),
        EditOperation.replace_para(
            op_id="op_replace", target_id=target,
            items=["第 1 条：责任条款 1 已被整段替换。"],
        ),
        EditOperation.delete_para(op_id="op_delete", target_ids=[stale]),
        EditOperation.replace_text(
            op_id="op_text", paragraph_id=anchor,
            find="Test Contract", replacement="Sample Contract",
        ),
    ])
    for op_result in result.results:
        print(op_result.op_id, op_result.op, op_result.status, op_result.new_ids)

    save = doc.save("contract.edited.docx")
    print("saved:", save.output_path, "backup:", save.backup_path)
```

Output on a first-time save (no prior destination, so no backup):

```text
op_insert insert_para_after success ('p_000029',)
op_replace replace_para success ('p_000030',)
op_delete delete_para success ()
op_text replace_text success ()
saved: contract.edited.docx backup: None
```

Every `OperationResult` correlates by input order and reports `new_ids` plus bounded previews for auditing.

## LLM-caller rules

- Only paragraph IDs returned by a read API in **this session** are valid; never invent IDs, XML, XPath, or offsets.
- Strings longer than ~40 chars, or any number, date, price, party name, URL, or email must go through `content_ref`.
- `raw=True` is forbidden.
- Do not call `save()` inside a batch, and never resubmit a failed batch verbatim.

See [Errors](errors.md) for the structured error contract and the recommended response per failure.
