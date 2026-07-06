## Why

Operating on document sections (heading → next heading) and performing global find-replace are the two most common agent workflows, yet both require multi-step boilerplate: grep for boundaries, collect IDs manually, loop over matches to build op lists. This friction leads to verbose scripts and increases the chance of off-by-one errors when identifying section ranges. Additionally, inserting content from one document into another loses all formatting because there's no paragraph-level copy path.

## What Changes

- Add `Document.replace_all(find, replacement, *, occurrence=-1)` — a convenience method that performs a global text replacement across all paragraphs in a single call, returning the total count of substitutions made.
- Add `Document.list_sections(heading_style_prefix="Heading")` — returns a list of section descriptors (heading paragraph ID, first body paragraph ID, last body paragraph ID) derived from paragraph style analysis.
- Add `Document.get_section(heading_id)` — returns the paragraph ID range for a section given its heading paragraph ID.
- Add `Document.copy_paragraphs_from(source_doc, start_id, end_id)` that extracts paragraph XML from one document and returns items suitable for `insert_para_before` / `insert_para_after` with `raw=True`, preserving run-level formatting.

## Capabilities

### New Capabilities
- `bulk-replace`: Document-wide find-and-replace convenience method that wraps per-paragraph `replace_text` operations internally.
- `section-ops`: Section-aware query APIs (`list_sections`, `get_section`) that expose heading-delimited ranges as first-class objects.
- `cross-doc-copy`: Cross-document paragraph copy that extracts raw OOXML paragraphs from a source document for insertion into a target document with formatting preserved.

### Modified Capabilities

(none)

## Impact

- **Code**: New public methods on `Document` class in `docx_knife/document.py`; potential helper module for section detection.
- **APIs**: Additive only — no breaking changes to existing interfaces.
- **SKILL.md**: Needs update to document the new convenience methods for agent use.
- **Tests**: New test cases for each capability.
