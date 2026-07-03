## ADDED Requirements

### Requirement: Ordered paragraph insertion
`insert_para_before` and `insert_para_after` SHALL accept a target ID and a nonempty ordered `items` array, flatten each item in item and paragraph order, and insert all results in that order. After-insertion MUST use a moving cursor or equivalent algorithm so repeated insertion cannot reverse results.

#### Scenario: Ordered insertion after anchor
- **WHEN** items expanding to `B` and `C` are inserted after `A` before `D`
- **THEN** document order is exactly `A, B, C, D`

#### Scenario: Ordered insertion before anchor
- **WHEN** items expanding to `B` and `C` are inserted before `D`
- **THEN** document order is exactly `B, C, D`

### Requirement: One-to-many paragraph replacement
`replace_para` SHALL require one target ID and a nonempty items array, replace the target with all expanded paragraphs in order, invalidate the target ID, and assign each result a new ID.

#### Scenario: Replace one paragraph with many
- **WHEN** two ordered paragraph values replace one target
- **THEN** both appear at the original position in input order, each has a new ID, and the old ID raises `ParagraphNotFoundError`

### Requirement: Multi-target paragraph deletion
`delete_para` SHALL require a nonempty, duplicate-free `target_ids` collection of existing IDs, remove all corresponding nodes and manifest entries, and execute physical removals in reverse document order.

#### Scenario: Invalid deletion target set
- **WHEN** target IDs are empty, duplicated, or include a nonexistent ID
- **THEN** prevalidation rejects the whole batch before any node or manifest entry is removed

### Requirement: Fluent paragraph object API
Document-level and `Paragraph` object APIs SHALL correspond to batch paragraph operations. Insert and replacement methods MUST return `list[Paragraph]` in document order; returned paragraphs MUST be immediately usable as anchors. `Paragraph.replace_para` and `Paragraph.delete_para` SHALL be single-target conveniences delegated to document-level behavior, and `Paragraph.read(raw=False)` SHALL support visible and raw reads.

#### Scenario: Chained paragraph construction
- **WHEN** a caller inserts a paragraph, then inserts after the returned object, then inserts before that second object
- **THEN** every returned object remains a valid anchor and final order reflects the three calls

### Requirement: Visible-mode paragraph construction
In visible mode, paragraph operations SHALL accept `list[str]`, expand paragraph boundaries defined by content processing, copy the anchor's `w:pPr`, and copy the first ordinary text run's `w:rPr` when present. If no ordinary text run exists, the system SHALL create a run without `w:rPr`; leading or trailing spaces MUST set `xml:space="preserve"`.

#### Scenario: Inserted paragraph inherits anchor formatting
- **WHEN** visible text is inserted beside an anchor with paragraph properties and a formatted ordinary text run
- **THEN** the new paragraph receives equivalent paragraph and run properties and preserves boundary spaces

### Requirement: Raw paragraph fragment mode
In raw mode, every item SHALL contain one or more complete, well-formed WordprocessingML `w:p` top-level elements. The system MUST parse and validate the entire fragment before DOM modification, preserve top-level order, allow arbitrary paragraph-internal XML, assign each top-level paragraph a new ID, and MUST NOT copy anchor properties, construct runs, project reserved markers, or normalize text. Non-`w:p` top-level nodes MUST raise `InvalidContentError`.

#### Scenario: Multi-paragraph raw replacement
- **WHEN** one raw item contains two well-formed `w:p` elements with custom internal XML
- **THEN** both replace the target in fragment order, preserve their supplied internals, and receive independent IDs

#### Scenario: Invalid raw fragment
- **WHEN** a raw fragment is malformed or has a non-`w:p` top-level element
- **THEN** validation fails before any DOM or manifest mutation

### Requirement: Content-mode consistency
All paragraph content operations and their returned previews SHALL use one `raw` mode per operation, defaulting to visible mode. Text and raw XML content MUST NOT be mixed within an operation. APIs without content MUST omit the parameter.

#### Scenario: Mixed content modes
- **WHEN** one operation combines visible text and raw XML items or mismatches result preview mode
- **THEN** schema validation rejects the operation before execution

### Requirement: Destructive replacement warnings
Whole-paragraph replacement SHALL execute even when it removes local run styling, hyperlinks, fields, bookmarks, comment ranges, or revision markup, but MUST list every detected removed structure in both `OperationResult.warnings` and the change log.

#### Scenario: Replace structured paragraph
- **WHEN** a target contains a hyperlink and local bold run and is replaced wholesale
- **THEN** replacement succeeds and warnings and change log identify the removed structures

### Requirement: Editable table paragraphs
Paragraph insertion, replacement, deletion, and object APIs SHALL operate identically on editable paragraphs in table cells, while preserving table containment and using IDs rather than row or column metadata for targeting.

#### Scenario: Edit nested-table paragraph
- **WHEN** a valid ID addresses a paragraph in a nested table cell
- **THEN** the requested edit occurs within that cell without using table coordinates as an execution key

