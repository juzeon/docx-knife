## Context

docx-knife exposes a paragraph-ID-centric editing model. All mutations go through `batch_edit` with explicit paragraph IDs. The `Document` class already provides `list_paragraphs`, `grep_paragraphs`, and per-paragraph text ops. Each `_ParagraphRecord` tracks `style_id` (the `w:pStyle` value), which is exactly what heading detection needs.

Currently, performing a document-wide find-replace requires: grep all matching paragraphs → build N `replace_text` ops → submit batch. Operating on a "section" (heading → next heading of same/higher level) requires the agent to manually walk paragraph IDs to find boundaries. Copying formatted paragraphs between documents has no supported path.

## Goals / Non-Goals

**Goals:**
- Provide `replace_all` as a single-call document-wide replacement that handles arbitrary match counts, returning a total substitution count.
- Provide `list_sections` / `get_section` that expose heading-delimited paragraph ranges using style-based heading detection.
- Provide `copy_paragraphs_from` that serializes paragraph XML from one `Document` and produces raw items usable with `insert_para_before`/`insert_para_after` with `raw=True`, preserving run properties.

**Non-Goals:**
- Heading detection from outline level or `w:numPr` (style-id prefix matching is sufficient for typical Chinese/English contracts).
- Cross-document style definition merging (styles are assumed compatible or manually reconciled).
- TOC-awareness or TOC update/regeneration.
- Nested section hierarchy (flat list of top-level sections is the v1 target).

## Decisions

### 1. `replace_all` wraps internal batch logic

`replace_all(find, replacement, *, regex=False, normalize_text=False)` iterates all paragraphs internally (using the existing index), matches within each paragraph's visible text, and performs replacements via the same `_apply_replace_text` code path that `batch_edit` uses. It does **not** go through `batch_edit` externally — instead it operates directly on the internal paragraph objects in a single pass, then invalidates the index.

**Rationale:** Avoids the 50-op batch limit and the overhead of constructing `EditOperation` objects for potentially hundreds of paragraphs. The internal code path already handles rollback at the paragraph level.

**Alternative considered:** Exposing a synthetic "global" op type in `EditOperation`. Rejected because it conflates the atomic-batch semantics (which are designed for targeted, composable edits) with a bulk sweep.

### 2. Section detection via style-id prefix matching

`list_sections(level=1)` scans the paragraph index for records whose `style_id` starts with `"Heading"` (or CJK equivalents like `"heading"`, case-insensitive). Each heading paragraph starts a section that extends until the next heading of the same or higher level (lower number), or end of document.

Returns `list[SectionInfo]` where `SectionInfo` contains:
- `heading_id: str` — the heading paragraph ID
- `heading_text: str` — visible text of the heading paragraph
- `level: int` — heading level (extracted from style suffix, e.g. "Heading1" → 1)
- `first_body_id: str | None` — first non-heading paragraph ID after heading (None if section is empty)
- `last_body_id: str | None` — last paragraph ID before the next section starts
- `paragraph_ids: tuple[str, ...]` — all paragraph IDs in the section (heading + body)

`get_section(heading_id)` returns a single `SectionInfo` for the given heading paragraph.

**Rationale:** Style-id is already indexed in `_ParagraphRecord.style_id`. No additional XML parsing needed. Covers the common case where headings use built-in styles.

**Alternative considered:** Full outline-level detection via `w:outlineLvl`. More accurate for edge cases but requires parsing `<w:pPr>` of every paragraph, and most docx files use named heading styles anyway.

### 3. Cross-document paragraph copy via raw XML extraction

`copy_paragraphs_from(source_doc, start_id, end_id)` extracts the XML of paragraphs from `start_id` to `end_id` (inclusive) in `source_doc`, serializes each `<w:p>` element to a string, and returns them as a `list[str]` suitable for passing to `insert_para_before`/`insert_para_after` with `raw=True`.

The method lives on the target `Document` instance: `target_doc.copy_paragraphs_from(source, ...)` — but since it just returns strings, it could also be a standalone function. We put it on `Document` for discoverability.

**Rationale:** Reuses the existing `raw=True` pathway which already handles namespace declarations and paragraph insertion. No new insertion logic needed.

**Caveat:** Style definitions are not copied. If the source uses styles not present in the target, formatting may degrade. This is an acceptable trade-off for v1 — users can manually ensure style compatibility or rely on Word's style fallback behavior.

## Risks / Trade-offs

- **[Performance of replace_all on large documents]** → Mitigated by operating directly on the in-memory lxml tree without intermediate data structures. The grep + replace is O(n) in paragraph count regardless.
- **[Heading style naming variations]** → Mitigated by case-insensitive matching and allowing a custom `style_prefix` parameter. Documents using non-standard heading styles would need to pass the prefix explicitly.
- **[Cross-doc style mismatch]** → Accepted trade-off for v1. Documented as a known limitation. The paragraph content and run properties are preserved; only style references may not resolve in the target.
- **[Section detection for flat headings only]** → v1 only detects top-level section boundaries. Nested sub-sections can be added later by filtering on level ranges.
