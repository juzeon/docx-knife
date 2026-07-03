# Quickstart

A 5-minute walkthrough that opens a DOCX, lists its paragraphs, submits a
single atomic batch that mixes an insert, a replace, a delete, and a
paragraph-internal `replace_text`, then saves.

## 1. Install

```bash
pip install docx-knife
```

## 2. Open a document

`Document.open` extracts `word/document.xml` into a private temp
workspace, captures a source fingerprint (used later to detect drift on
save), and returns a live [`Document`](api.md#docx_knife.Document) object.
Use it as a context manager so the temp workspace is cleaned up on exit.

```python
from docx_knife import Document

with Document.open("contract.docx") as doc:
    print(f"paragraphs: {doc.paragraph_count()}")
```

## 3. Discover paragraph IDs

Never invent IDs. Every batch operation must reference IDs returned by
one of the read APIs inside the current session.

```python
from docx_knife import Document

with Document.open("contract.docx") as doc:
    result = doc.list_paragraphs(start=1, limit=5, max_chars=40)
    for info in result.paragraphs:
        print(info.id, "|", info.text)
```

Example output (with the fixture used in the test suite):

```text
p_000001 | 契约合同 — Test Contract
p_000002 | 第一部分 定义
p_000003 | 第 1 条：定义条款 1 的正文内容。
p_000004 | 第 2 条：定义条款 2 的正文内容。
p_000005 | 第 3 条：定义条款 3 的正文内容。
```

Use [`grep_paragraphs`](api.md#docx_knife.Document) to locate a specific
paragraph:

```python
hits = doc.grep_paragraphs("target", regex=False, max_chars=60)
for match in hits.matches:
    print(match.paragraph.id, match.match_count, "|", match.paragraph.text)
```

## 4. Build one atomic batch

Assemble every edit into a single call to
[`batch_edit`](api.md#docx_knife.Document). One `batch_edit` is one
transaction: on any failure the DOM is rolled back exactly, so partial
success is impossible.

```python
from docx_knife import Document, EditOperation

with Document.open("contract.docx") as doc:
    # First: locate the IDs you actually want to touch.
    target = doc.grep_paragraphs("责任条款 1").matches[0].paragraph.id
    anchor = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    stale = doc.list_paragraphs(start=doc.paragraph_count(), limit=1).paragraphs[0].id

    batch = [
        EditOperation.insert_para_after(
            op_id="op_insert",
            target_id=anchor,
            items=["This clause was inserted by docx-knife."],
        ),
        EditOperation.replace_para(
            op_id="op_replace",
            target_id=target,
            items=["第 1 条：责任条款 1 已被整段替换。"],
        ),
        EditOperation.delete_para(
            op_id="op_delete",
            target_ids=[stale],
        ),
        EditOperation.replace_text(
            op_id="op_text",
            paragraph_id=anchor,
            find="Test Contract",
            replacement="Sample Contract",
        ),
    ]

    result = doc.batch_edit(batch)
    for op_result in result.results:
        print(op_result.op_id, op_result.op, op_result.status, op_result.new_ids)

    save = doc.save("contract.edited.docx")
    print("saved:", save.output_path, "backup:", save.backup_path)
```

Example output on a first-time save (no prior destination, so no backup
is produced):

```text
op_insert insert_para_after success ('p_000029',)
op_replace replace_para success ('p_000030',)
op_delete delete_para success ()
op_text replace_text success ()
saved: /abs/path/contract.edited.docx backup: None
```

Every [`OperationResult`](api.md#docx_knife.OperationResult) is correlated
by input order and reports `new_ids` when paragraphs were created, plus
bounded before/after previews for auditing.

## Rules for LLM callers

These are the same prohibitions enforced by
[`skills/docx-knife/SKILL.md`](https://github.com/anon/docx-knife/blob/main/skills/docx-knife/SKILL.md):

!!! warning "Prohibitions"
    - **Never emit** XML, XPath, array indexes, `w14:paraId`, character
      offsets, or any invented ID. Only IDs returned by the read APIs in
      the current session are valid.
    - **Never quote long source text back** into `content_literal`.
      Anything longer than about 40 characters, and any number, date,
      monetary amount, company name, URL, or email address, must use
      `content_ref` (jsonpath / file / command).
    - **Never use `raw=True`.** The LLM-facing schema forbids the `raw`
      field; raw XML is a trusted-caller-only channel.
    - **Never call `save()` in the middle of a batch.** One
      `batch_edit(...)` call is one atomic write; call
      `save(output_path)` once at the end.
    - **Never retry a failed batch verbatim.** Read the structured
      error, rebuild the batch (fresh IDs, corrected selectors), then
      submit again.

See [Errors](errors.md) for the structured error contract and the
recommended response for each failure.
