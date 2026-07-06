# Content sources and selectors

Four channels; exactly one of `content_literal` or `content_ref` per item, else [`InvalidContentError`](errors.md).

| Kind | Use when |
| --- | --- |
| `content_literal` | Short human phrase ≤ ~40 chars, no numbers/dates/prices/party names/URLs/emails/trademarks. |
| `content_ref` — `jsonpath` | Structured data from a canonical JSON source. |
| `content_ref` — `file` | Long clauses, boilerplate, or previously reviewed drafts on disk. |
| `content_ref` — `command` | Values produced by a deterministic script (formatter, template renderer, date normalizer, **cross-document extraction**). |

Default to `content_ref` for anything longer than ~40 chars or containing deterministic values.

**Hard rule**: text originating from another document, file, or any external source must **always** flow through `content_ref`. Never read external text into a variable and pass it as `content_literal` — this defeats the purpose of keeping large text out of the agent context and risks hallucination drift. When no pre-existing file is available, write a helper script and use `content_ref` — `command`.

## `content_literal`

```python
from docx_knife import EditOperation

EditOperation.insert_para_after(
    op_id="op1",
    target_id="p_000042",
    items=["Short human phrase."],
)
```

Normalized identically to every other source (see [Newline expansion](#newline-expansion-visible-mode)).

## `content_ref` — `jsonpath`

```python
from docx_knife import ContentItem, ContentSourceJsonPath, EditOperation

item = ContentItem(
    content_ref=ContentSourceJsonPath(source="contract.json", path="$.party_a.name"),
)
EditOperation.replace_text(op_id="op2", paragraph_id="p_000005",
                           find="Party A", replacement=item)
```

Must resolve to **exactly one** scalar (str/int/float/bool). Missing keys or multi-value hits → [`InvalidContentError`](errors.md). Source path must resolve inside `ContentResolverConfig.input_roots`.

## `content_ref` — `file`

```python
from docx_knife import ContentItem, ContentSourceFile

ContentItem(content_ref=ContentSourceFile(path="clauses/confidentiality.txt", encoding="utf-8"))
```

Path resolved with `strict=True` inside an allowed input root; symlink-based escape rejected. Decoding failures → [`InvalidContentError`](errors.md).

## `content_ref` — `command`

```python
from docx_knife import ContentItem, ContentSourceCommand

ContentItem(content_ref=ContentSourceCommand(
    argv=("python", "scripts/render_clause.py", "--contract", "contract.json"),
    timeout_seconds=30.0,
))
```

- `argv` runs without a shell. `cwd` defaults to the document workspace and must lie inside it if provided.
- `timeout_seconds` is mandatory (> 0). Timeout, non-zero exit, stdout > 1 MiB, or non-UTF-8 stdout → [`InvalidContentError`](errors.md).
- Only env vars in `command_env_allowlist` (default: `PATH`, `LANG`, `LC_ALL`) are inherited; extras go through the reference's `env`.

## Newline expansion (visible mode)

Applied to `content_literal`, jsonpath, file, and command results:

1. `\r\n` and `\r` → `\n`.
2. Single `\n` inside a paragraph → `<w:br/>`.
3. ≥ 2 consecutive `\n` → paragraph split; each new paragraph inherits `w:pPr` / first-run `w:rPr` from the anchor.

## `normalize_text` (opt-in)

`batch_edit(..., normalize_text=True)` (or the same flag on paragraph fluent methods) applies a deterministic pass to every resolved text before write:

- **Chinese punctuation.** Half-width `, ? ! : ; ( )` → full-width `，？！：；（）` when adjacent to CJK. `.` → `。` only when both sides are CJK (or terminating a CJK phrase); ASCII decimals and abbreviations preserved.
- **CJK / Latin spacing.** Single space between a CJK character and an adjacent ASCII letter/digit.
- **Byte-preserved substrings.** URLs matching `https?://\S+`, emails matching `\S+@\S+\.\S+`, and backtick-wrapped inline code spans are kept verbatim.
- Leading/trailing whitespace is never trimmed.

Off by default. Leave off for exact input preservation.

## Selectors (`find`)

- `regex=False` (default): literal substring. Empty patterns rejected.
- `regex=True`: anchored Python regex compiled once; compile errors → [`InvalidPatternError`](errors.md). Zero-length matches rejected.

Evaluated against final visible text (TextMap): `<w:t>` + `<w:ins>` runs, excluding `<w:del>`. Tabs, line/page/column breaks, and CRs are projected as reversible reserved markers (`[[DOCX:TAB]]`, `[[DOCX:LINE_BREAK]]`, `[[DOCX:PAGE_BREAK]]`, `[[DOCX:COLUMN_BREAK]]`, `[[DOCX:CR]]`). Matches crossing an atomic marker range → [`UnsupportedStructureError`](errors.md); use the marker literal to hit one explicitly.

## Raw mode

`raw=True` bypasses text expansion and normalization, feeding WordprocessingML fragments directly into the tree. Use it when you need exact OOXML control (custom `w:rPr`, styles, fields, structured document tags). Accepted on `insert_para_before`, `insert_para_after`, and `replace_para`; paragraph-internal ops (`replace_text`, `insert_text_before`, `insert_text_after`, `delete_text`) reject it unconditionally.

Raw items must each contain one or more top-level `<w:p>` elements in the standard WordprocessingML namespace. Mixed modes (raw + visible in one item) and content outside that shape are refused with [`InvalidContentError`](errors.md).

```python
from docx_knife import Document, EditOperation

fragment = (
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">附录 A</w:t></w:r>'
    '</w:p>'
)

with Document.open("contract.docx") as doc:
    anchor = doc.list_paragraphs(start=1, limit=1).paragraphs[0].id
    doc.batch_edit([
        EditOperation.insert_para_after(
            op_id="op1", target_id=anchor, items=[fragment], raw=True,
        ),
    ])
    doc.save("contract.edited.docx")
```
