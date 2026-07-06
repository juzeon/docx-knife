# Operations

Two families: **paragraph-level** (whole-paragraph insert / replace / delete) and **paragraph-internal** (span-level edits inside one paragraph, located by a text selector).

Common fields:

- `op_id` — unique inside the batch; correlates results.
- `target_id` — paragraph ID from a read API (single-target ops).
- `target_ids` — tuple of unique paragraph IDs (`delete_para` only).
- `items` — content items expanded into new paragraphs (paragraph inserts / replace). See [content-sources.md](content-sources.md).
- `find` — text selector (paragraph-internal ops). See [content-sources.md](content-sources.md#selectors-find).
- `occurrence` — 0-based match index, `-1` for all (right-to-left), or `None` to require a unique match.
- `content_literal` / `content_ref` — mutually exclusive (paragraph-internal insert/replace).
- `raw` — accepted on paragraph-level ops (`insert_para_before`, `insert_para_after`, `replace_para`) to submit exact `<w:p>` fragments; paragraph-internal ops reject it. See [Raw mode](content-sources.md#raw-mode).

## Overview

| Op | Target | Content mode | Produces new IDs? | Notes |
| --- | --- | --- | --- | --- |
| `insert_para_before` | one ID | items (visible or raw) | yes, in document order | inherits `w:pPr` / first ordinary `w:rPr` from anchor |
| `insert_para_after` | one ID | items (visible or raw) | yes, moving-cursor preserves item order |
| `replace_para` | one ID | items (visible or raw) | yes; old ID permanently invalidated | protected structures emitted as warnings |
| `delete_para` | many unique IDs | — | no | reverse-document-order removal |
| `replace_text` | one ID | `content_literal` xor `content_ref` (visible only) | no | run chain rebuilt; `<w:rPr>` inherited from left boundary |
| `insert_text_before` | one ID | `content_literal` xor `content_ref` (visible only) | no | spliced at left of match |
| `insert_text_after` | one ID | `content_literal` xor `content_ref` (visible only) | no | spliced at right of match |
| `delete_text` | one ID | — | no | matched span removed; atomic reserved-marker ranges respected |

## `insert_para_before` / `insert_para_after`

Insert one or more paragraphs immediately before / after `target_id`. `_after` uses a moving cursor so item order is preserved regardless of item count.

Visible items are normalized (`\r\n`/`\r` → `\n`; single `\n` → `<w:br/>`; ≥ 2 consecutive `\n` → paragraph split). Every new paragraph inherits `w:pPr` from the anchor and, if present, the anchor's first ordinary run `w:rPr`. Raw items must each contain one or more top-level `<w:p>` elements; internal order is verbatim.

New IDs are allocated in document order and returned as `OperationResult.new_ids`.

## `replace_para`

Replaces `target_id` with one or more new paragraphs; the old ID is permanently invalidated. Detected protected structures (`w:fldChar`, `w:hyperlink`, `w:bookmarkStart`, `w:ins`, `w:del`, ...) are emitted as warnings — `replace_para` does not refuse them.

## `delete_para`

Empty list, duplicated IDs, or unknown IDs raise [`ValidationError`](errors.md) (`failed_check` ∈ `nonempty`, `unique`, `resolvable`) before any DOM mutation. Nodes are removed in reverse document order; each ID is permanently invalidated.

## `replace_text`, `insert_text_before`, `insert_text_after`, `delete_text`

Selector runs against the paragraph's final visible text (TextMap: `<w:t>` and `<w:ins>` runs; `<w:del>` excluded; breaks and tabs projected as reversible reserved markers).

- Exactly one of `content_literal` / `content_ref` (insert/replace only). `raw=True` is rejected.
- Empty literal selector or a regex that fails to compile → [`InvalidPatternError`](errors.md).
- `occurrence is None`: exactly one match required. Multiple → [`AmbiguousTextMatchError`](errors.md); zero → [`TextNotFoundError`](errors.md). `occurrence=-1` executes every match, right-to-left, so earlier offsets remain valid.
- New content inherits `<w:rPr>` from the boundary it attaches to.
- A match that crosses an atomic marker (`w:tab`, `w:br`, `w:cr`) or a protected structure → [`UnsupportedStructureError`](errors.md).

`insert_text_*` splices around the match without modifying it. `delete_text` removes the span.

## Conflict matrix

Prevalidated across the whole batch; every violation raises a `ValidationError` at the `prevalidation` stage.

| Combination on the same target | Allowed? |
| --- | --- |
| Two `insert_para_before` on the same anchor | yes, merged left-to-right in input order |
| Two `insert_para_after` on the same anchor | yes, merged in input order (moving cursor) |
| `insert_para_before` + `insert_para_after` on the same anchor | yes |
| `replace_para` on the same target | at most one |
| `delete_para` targeting an ID another op references | no |
| `replace_para` + text op on the same target | no |
| Two text ops touching the same character range on the same paragraph | no |
| Any op referencing an ID a previous op invalidated | no |
| Duplicate IDs in one `delete_para.target_ids` | no |
| Duplicate `op_id` in the batch | no |
