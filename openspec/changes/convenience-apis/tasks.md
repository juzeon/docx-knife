## 1. Data Model

- [x] 1.1 Add `SectionInfo` dataclass to `docx_knife/_models.py` with fields: `heading_id`, `heading_text`, `level`, `body_ids`, `all_ids`
- [x] 1.2 Export `SectionInfo` from `docx_knife/__init__.py`

## 2. Bulk Replace

- [x] 2.1 Implement `Document.replace_all(find, replacement, *, regex=False, normalize_text=False)` in `docx_knife/document.py` using internal TextMap-based replacement logic (same as `_apply_replace_text`), returning total substitution count
- [x] 2.2 Add tests for `replace_all`: literal match across multiple paragraphs, regex match, zero-match case, cross-run match, formatting preservation

## 3. Section Operations

- [x] 3.1 Add helper function `_detect_heading_level(style_id, style_prefix)` that returns the heading level (int) or None if not a heading
- [x] 3.2 Implement `Document.list_sections(*, level=None, style_prefix="Heading")` that scans `_ParagraphRecord` entries and builds `SectionInfo` objects
- [x] 3.3 Implement `Document.get_section(heading_id)` that returns the `SectionInfo` for a specific heading paragraph, raising `ValueError` for non-headings and `ParagraphNotFoundError` for invalid IDs
- [x] 3.4 Add tests for `list_sections`: multiple headings, level filtering, empty document, consecutive headings, nested headings included in parent
- [x] 3.5 Add tests for `get_section`: valid heading, non-heading ID, invalid ID

## 4. Cross-Document Copy

- [x] 4.1 Implement `Document.copy_paragraphs_from(source, start_id, end_id)` that serializes `<w:p>` elements from source document with proper namespace declarations
- [x] 4.2 Add tests for `copy_paragraphs_from`: range copy with formatting preserved, single paragraph, invalid IDs, reversed range order, namespace presence

## 5. Integration & Documentation

- [x] 5.1 Update `skills/docx-knife/SKILL.md` with usage examples for `replace_all`, `list_sections`, `get_section`, and `copy_paragraphs_from`
- [x] 5.2 Run full test suite to verify no regressions
