# Content sources and selectors

Every paragraph-level insert / replace and every paragraph-internal
insert / replace can source its text from one of four channels:

1. `content_literal` — an inline UTF-8 string.
2. `content_ref` with `type=jsonpath` — a single scalar from a JSON file.
3. `content_ref` with `type=file` — a UTF-8 (or specified-encoding) text
   file.
4. `content_ref` with `type=command` — stdout of a subprocess launched
   with `argv` (no shell).

Exactly one of `content_literal` or `content_ref` may be provided per
item; violating that cardinality raises
[`InvalidContentError`](errors.md).

## When to use which

| Kind | Use when |
| --- | --- |
| `content_literal` | Short human phrases that the model can safely generate: ≤ ~40 characters and free of numbers, dates, prices, party names, URLs, emails, or trademarks. |
| `jsonpath` | Structured data extracted from a canonical JSON source. |
| `file` | Long clauses, boilerplate, or previously reviewed drafts checked into a workspace. |
| `command` | Values produced by a deterministic script (formatter, template renderer, date/time normalizer). |

LLM callers should default to `content_ref` for anything longer than
~40 characters or containing deterministic values.

## `content_literal`

```python
from docx_knife import EditOperation

EditOperation.insert_para_after(
    op_id="op1",
    target_id="p_000042",
    items=["Short human phrase."],
)
```

The literal is normalized identically to every other source (CRLF/CR
collapse, single `\n` → `<w:br/>`, two or more `\n` → paragraph
boundary) so one item can expand into several paragraphs.

## `content_ref` — jsonpath

```python
from docx_knife import (
    ContentItem,
    ContentSourceJsonPath,
    EditOperation,
)

item = ContentItem(
    content_ref=ContentSourceJsonPath(
        source="contract.json",
        path="$.party_a.name",
    ),
)
EditOperation.replace_text(
    op_id="op2",
    paragraph_id="p_000005",
    find="Party A",
    replacement=item,
)
```

- Must resolve to **exactly one** scalar (string, int, float, or bool).
- Missing keys or multi-value hits raise
  [`InvalidContentError`](errors.md).
- Source paths must resolve inside the configured
  `ContentResolverConfig.input_roots`.

## `content_ref` — file

```python
from docx_knife import ContentItem, ContentSourceFile

ContentItem(
    content_ref=ContentSourceFile(
        path="clauses/confidentiality.txt",
        encoding="utf-8",
    ),
)
```

- Path is resolved with `strict=True` and must be inside an allowed
  input root; symlink-based escape is rejected.
- Decoding failures (bad encoding, invalid bytes) raise
  [`InvalidContentError`](errors.md).

## `content_ref` — command

```python
from docx_knife import ContentItem, ContentSourceCommand

ContentItem(
    content_ref=ContentSourceCommand(
        argv=("python", "scripts/render_clause.py", "--contract", "contract.json"),
        timeout_seconds=30.0,
    ),
)
```

- `argv` is required and passed through without a shell. The working
  directory is the document's private workspace unless an explicit
  `cwd` is supplied and lies inside the workspace.
- `timeout_seconds` is mandatory; a timeout, non-zero exit code, stdout
  larger than the internal 1 MiB limit, or non-UTF-8 stdout all raise
  [`InvalidContentError`](errors.md).
- Only environment variables in the resolver's
  `command_env_allowlist` (default: `PATH`, `LANG`, `LC_ALL`) are
  inherited from the parent; extra values must be provided via the
  reference's `env` mapping.

## Newline expansion (visible mode)

Applied identically to `content_literal`, jsonpath, file, and command
results:

1. `\r\n` and `\r` are normalized to `\n`.
2. A single `\n` inside a paragraph becomes a `<w:br/>` line break.
3. Two or more consecutive `\n` split the string into separate
   paragraphs; each new paragraph inherits `w:pPr` / first-run `w:rPr`
   from the anchor.

## `normalize_text` rules

`batch_edit(..., normalize_text=True)` (or the same flag on the
paragraph fluent methods) applies a deterministic normalization pass to
every resolved text before it is written:

- **Chinese punctuation.** Half-width `, ? ! : ; ( )` become full-width
  `\uff0c \uff1f \uff01 \uff1a \uff1b \uff08 \uff09` when adjacent to a
  CJK character. `.` becomes `\u3002` only when both sides are CJK (or
  it terminates a CJK phrase); embedded ASCII decimals and abbreviations
  are preserved.
- **CJK / Latin spacing.** Inserts a single space between a CJK
  character and an adjacent ASCII letter or digit.
- **Byte-for-byte protected substrings.** URLs matching
  `https?://\S+`, email addresses matching `\S+@\S+\.\S+`, and inline
  code spans wrapped in backticks are preserved verbatim, including
  any punctuation they contain.
- Leading and trailing whitespace is never trimmed.

Off by default. Leave it off if you want exact input preservation.

## Selectors (`find`)

Every paragraph-internal op locates its target span with a `Selector`:

- `regex=False` (default): a literal substring. Empty patterns are
  rejected.
- `regex=True`: an anchored Python regex compiled once; compile errors
  raise [`InvalidPatternError`](errors.md).

Selectors are evaluated against the paragraph's final visible text
(TextMap): text from `<w:t>` and inside `<w:ins>` runs, but not
`<w:del>` runs. Tabs, line/page/column breaks, and carriage returns are
projected as reversible reserved markers (`[[DOCX:TAB]]`,
`[[DOCX:LINE_BREAK]]`, `[[DOCX:PAGE_BREAK]]`, `[[DOCX:COLUMN_BREAK]]`,
`[[DOCX:CR]]`). A match that crosses a marker's atomic range raises
[`UnsupportedStructureError`](errors.md); use the marker literal in
your pattern to hit it explicitly.

Regex zero-length matches are rejected.

## Raw mode

!!! danger "Trusted callers only"
    `raw=True` bypasses text expansion and normalization and feeds
    supplied WordprocessingML fragments directly into the tree. The
    LLM-facing JSON schema forbids this field; use it only from
    trusted programmatic callers, and only for paragraph-level insert /
    replace operations. All paragraph-internal ops
    (`replace_text`, `insert_text_before`, `insert_text_after`,
    `delete_text`) reject `raw=True` unconditionally.

Raw items must each contain one or more top-level `<w:p>` elements in
the standard WordprocessingML namespace; the parser refuses mixed modes
(raw + visible in the same item) and any content outside that shape.
