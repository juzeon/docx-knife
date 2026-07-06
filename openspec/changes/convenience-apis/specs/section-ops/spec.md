## ADDED Requirements

### Requirement: List document sections
The `Document` class SHALL provide a `list_sections(*, level=None, style_prefix="Heading")` method that returns a list of `SectionInfo` objects representing contiguous paragraph ranges delimited by heading-style paragraphs.

#### Scenario: List all top-level sections
- **WHEN** `doc.list_sections()` is called on a document with headings styled "Heading1", "Heading2", etc.
- **THEN** a list of `SectionInfo` objects is returned, one per heading paragraph, ordered by document position

#### Scenario: Filter sections by level
- **WHEN** `doc.list_sections(level=1)` is called
- **THEN** only sections whose heading style corresponds to level 1 (e.g., "Heading1") are returned

#### Scenario: Document with no headings
- **WHEN** `doc.list_sections()` is called on a document with no heading-styled paragraphs
- **THEN** an empty list is returned

### Requirement: SectionInfo structure
Each `SectionInfo` object SHALL expose:
- `heading_id: str` — paragraph ID of the heading
- `heading_text: str` — visible text of the heading paragraph
- `level: int` — heading level number (extracted from style name suffix)
- `body_ids: tuple[str, ...]` — paragraph IDs of all non-heading paragraphs in the section (between this heading and the next heading of same or higher level)
- `all_ids: tuple[str, ...]` — all paragraph IDs including the heading

#### Scenario: Section contains body paragraphs
- **WHEN** a heading at level 1 is followed by 5 body paragraphs before the next level-1 heading
- **THEN** `section.body_ids` contains the 5 body paragraph IDs and `section.all_ids` has 6 entries (heading + 5 body)

#### Scenario: Empty section (consecutive headings)
- **WHEN** two headings appear consecutively with no body paragraphs between them
- **THEN** the first heading's `SectionInfo` has an empty `body_ids` tuple

### Requirement: Get a single section by heading ID
The `Document` class SHALL provide a `get_section(heading_id)` method that returns the `SectionInfo` for the section starting at the given heading paragraph.

#### Scenario: Valid heading ID
- **WHEN** `doc.get_section(heading_id)` is called with a valid heading paragraph ID
- **THEN** the corresponding `SectionInfo` is returned

#### Scenario: Non-heading paragraph ID
- **WHEN** `doc.get_section(body_paragraph_id)` is called with an ID that is not a heading paragraph
- **THEN** a `ValueError` is raised indicating the paragraph is not a heading

#### Scenario: Invalid paragraph ID
- **WHEN** `doc.get_section("p_nonexistent")` is called
- **THEN** a `ParagraphNotFoundError` is raised

### Requirement: Section boundary detection
Section boundaries SHALL be determined by heading level: a section extends from its heading paragraph until (exclusive) the next paragraph whose heading level is less than or equal to the current section's level, or until end of document.

#### Scenario: Nested headings are included in parent section body
- **WHEN** a Heading1 is followed by a Heading2, then body text, then another Heading1
- **THEN** the first Heading1's section includes the Heading2 and its body text in `body_ids`; the section ends just before the second Heading1

#### Scenario: Same-level heading terminates section
- **WHEN** two consecutive Heading1 paragraphs appear
- **THEN** the first section ends just before the second Heading1
