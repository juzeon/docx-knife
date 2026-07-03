# Operations

Every mutation is expressed as one operation. Operations come in two
families: **paragraph-level** (whole-paragraph insert / replace / delete)
and **paragraph-internal** (span-level edits inside a single paragraph
found by a text selector).

All operations share these fields:

- `op_id` — unique string inside the batch; used to correlate results.
- `target_id` — a paragraph ID from a read API (single-target ops).
- `target_ids` — tuple of unique paragraph IDs (`delete_para` only).
- `items` — content items expanded into new paragraphs (paragraph
  inserts / replace). See [Content sources](content-sources.md).
- `find` — text selector (paragraph-internal ops). See
  [Content sources](content-sources.md#selectors-find).
- `occurrence` — 0-based index into all matches, `-1` for every match
  (right-to-left execution), or `None` to require a unique match.
- `content_literal` / `content_ref` — mutually exclusive content source
  (paragraph-internal insert/replace).
- `raw` — **trusted-caller only.** Rejected by the JSON schema.
  Paragraph-internal ops reject `raw=True` unconditionally.

## Overview

| Op | Target | Content mode | Produces new paragraphs? | Notes |
| --- | --- | --- | --- | --- |
| `insert_para_before` | one paragraph ID | items (visible or raw) | yes, in document order | inherits `w:pPr` / first ordinary `w:rPr` from anchor when visible |
| `insert_para_after` | one paragraph ID | items (visible or raw) | yes, in document order | moving-cursor insert preserves item order |
| `replace_para` | one paragraph ID | items (visible or raw) | yes, old ID is invalidated | detected protected structures emitted as warnings |
| `delete_para` | many unique IDs | — | no | fails fast on empty / duplicated / unknown IDs; reverse-document-order removal |
| `replace_text` | one paragraph ID | `content_literal` xor `content_ref` (visible only) | no | rebuilds run chain; run properties inherited from left boundary |
| `insert_text_before` | one paragraph ID | `content_literal` xor `content_ref` (visible only) | no | inserted at the left of a match boundary |
| `insert_text_after` | one paragraph ID | `content_literal` xor `content_ref` (visible only) | no | inserted at the right of a match boundary |
| `delete_text` | one paragraph ID | — | no | removes the matched span; reserved-marker atomic ranges are respected |

## `insert_para_before`

Inserts one or more paragraphs immediately before `target_id`.

- Items are resolved to paragraph texts. Visible-mode items are
  normalized (`\r\n`/`\r` → `\n`), single `\n` becomes `<w:br/>`, and
  two or more consecutive `\n` split into paragraph boundaries.
- Every new paragraph inherits `w:pPr` from the anchor and, if the first
  ordinary run of the anchor has one, its `w:rPr`.
- Raw items must each contain one or more top-level WordprocessingML
  `<w:p>` elements; internal ordering is preserved verbatim.

New IDs are allocated in document order and returned as the
`OperationResult.new_ids` tuple.

## `insert_para_after`

Same semantics as [`insert_para_before`](#insert_para_before) but each
new paragraph is placed after the previous cursor position, starting at
`target_id`. Item order is preserved regardless of item count.

## `replace_para`

Replaces the single paragraph `target_id` with one or more new
paragraphs.

- The old ID is invalidated permanently. Any subsequent reference to it
  raises [`ParagraphNotFoundError`](errors.md).
- Every protected structure detected inside the removed paragraph
  (`w:fldChar`, `w:hyperlink`, `w:bookmarkStart`, `w:ins`, `w:del`, ...)
  is emitted as a warning. `replace_para` does **not** refuse the edit
  when structures are present — it is a wholesale replacement — but the
  warnings make the loss auditable.
- New IDs are allocated in document order.

## `delete_para`

Deletes every paragraph in `target_ids`.

- Prevalidation: empty list, duplicated IDs, and unknown IDs each cause
  a [`ValidationError`](errors.md) (`failed_check` ∈ `nonempty`,
  `unique`, `resolvable`) before any DOM mutation.
- Nodes are removed in reverse document order so earlier deletions
  cannot shift later indexes.
- Each deleted ID is invalidated permanently.

## `replace_text`

Runs a selector against the paragraph's final visible text (TextMap:
`<w:t>` and `<w:ins>` runs included, `<w:del>` runs excluded, breaks and
tabs projected as reversible reserved markers), and replaces the matched
span with the resolved content.

- Content mode: exactly one of `content_literal` or `content_ref`.
- `raw=True` is rejected: paragraph-internal ops never accept raw XML.
- Selector is compiled once. Literal selectors must be non-empty; regex
  selectors that fail to compile raise
  [`InvalidPatternError`](errors.md).
- If `occurrence is None`, exactly one match is required. Multiple
  matches raise [`AmbiguousTextMatchError`](errors.md); zero matches
  raise [`TextNotFoundError`](errors.md).
- `occurrence=-1` executes every match, right-to-left, so earlier char
  offsets remain valid.
- Replacement inherits `<w:rPr>` from the left boundary of the match.
- A match that crosses an atomic marker (`w:tab`, `w:br`, `w:cr`) or a
  protected structure raises
  [`UnsupportedStructureError`](errors.md).

## `insert_text_before` / `insert_text_after`

Same match resolution as `replace_text`, but the new content is spliced
immediately before (or after) the matched span rather than replacing it.

- Inserted run inherits `<w:rPr>` from the boundary position it is
  attached to.
- The match itself is not modified.

## `delete_text`

Same match resolution as `replace_text`, but the matched span is removed
and no replacement is spliced in.

## Conflict matrix

`batch_edit` enforces the following conflict rules across the whole
batch before execution begins. Every violation raises a structured
`ValidationError` at the `prevalidation` stage:

| Combination on the same target | Allowed? |
| --- | --- |
| Two `insert_para_before` on the same anchor | yes, merged left-to-right in input order |
| Two `insert_para_after` on the same anchor | yes, merged in input order (moving cursor) |
| `insert_para_before` + `insert_para_after` on the same anchor | yes |
| `replace_para` on the same target | at most one |
| `delete_para` referencing an ID that another op targets | no |
| `replace_para` + text op on the same target | no (target is destroyed) |
| Two text ops touching the same character range on the same paragraph | no |
| Any op referencing an ID that a previous op already invalidated | no |
| Duplicate IDs inside a single `delete_para.target_ids` | no |
| Two operations sharing the same `op_id` | no |

The full matrix is executed as a prevalidation pass, so no operation
runs until the whole batch is proven consistent.
