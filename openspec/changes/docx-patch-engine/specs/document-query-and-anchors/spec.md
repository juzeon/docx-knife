## ADDED Requirements

### Requirement: Editable paragraph discovery
The system SHALL discover editable `w:p` elements in document order from the body and table cells of `word/document.xml`, including nested tables, and MUST exclude every paragraph below an SDT. Excluded paragraphs MUST retain their structural contribution to global numbering and table-position calculations. Headers, footers, notes, comments, text boxes, and other parts MUST remain outside the query and editing surface.

#### Scenario: SDT content is excluded without collapsing position
- **WHEN** a document contains ordinary paragraphs and an SDT descendant paragraph
- **THEN** the SDT paragraph receives no ID and is absent from query results, while later paragraphs retain global and table positions that account for the excluded structure

### Requirement: Instance-local paragraph anchors
On open, the system SHALL assign each editable paragraph a unique ID such as `p_000001` in document order and bind it to its live XML node. IDs MUST be execution coordinates only within the current `Document` instance; `w14:paraId` and indexes MUST be diagnostic metadata and MUST NOT be used for write targeting or fuzzy relocation.

#### Scenario: Missing and duplicate native IDs
- **WHEN** editable paragraphs have missing or duplicate `w14:paraId` values
- **THEN** every editable paragraph receives a distinct agent-owned ID and remains addressable by that ID

#### Scenario: Detached anchor is rejected
- **WHEN** an operation references an ID whose node has been deleted or detached from the current tree
- **THEN** the system raises `ParagraphNotFoundError` and does not relocate the target by text, index, or similarity

### Requirement: Anchor lifecycle
Paragraph text edits SHALL preserve the paragraph ID and node binding. Inserted and replacement paragraphs SHALL receive new monotonically increasing IDs; replaced and deleted IDs MUST be removed from the manifest and MUST never be reused. Reopening or restoring a process SHALL parse the XML again and issue new instance-local IDs.

#### Scenario: Replacement invalidates the old ID
- **WHEN** a paragraph is replaced by two paragraphs
- **THEN** the old ID becomes invalid and each replacement receives a new ID in document order

### Requirement: Paragraph listing and pagination
The system SHALL expose `paragraph_count()` and `list_paragraphs(start=1, limit=None, max_chars=80, raw=False)`. A list result MUST include paragraphs and pagination metadata; each paragraph MUST include its ID, global ordinal, style, structural location, and either visible `text` or raw `xml`. Positive `max_chars` SHALL truncate each returned value, while `max_chars <= 0` SHALL return it in full. Pagination and truncation MUST NOT alter IDs or global ordinals.

#### Scenario: Paginated listing preserves identity
- **WHEN** a caller requests a limited window with positive `max_chars`
- **THEN** only that window is returned, values are bounded, and IDs and global ordinals equal those in an unpaginated listing

### Requirement: Paragraph search
The system SHALL expose `grep_paragraphs(pattern, regex=False, start=1, limit=None, max_chars=0, raw=False)`. Search MUST inspect each complete paragraph value before truncating results, use literal matching when `regex=false`, and use regular expressions when `regex=true`. The result MUST include matching paragraphs, IDs, ranges in the complete searched value, total match count, and listing-compatible pagination metadata.

#### Scenario: Match lies beyond returned preview
- **WHEN** a complete paragraph matches but the match begins after the requested `max_chars` boundary
- **THEN** the paragraph is returned with the complete-value match range and a truncated display value

### Requirement: Paragraph and document reads
The system SHALL expose `get_paragraph(paragraph_id, raw=False)` and `get_visible_text(raw=False)`. In visible mode the methods MUST use visible-text projection. In raw mode a paragraph read MUST return its complete `w:p` element, and document text MUST concatenate complete paragraph XML in document order without an XML declaration or wrapper.

#### Scenario: Raw document read
- **WHEN** `get_visible_text(raw=true)` is called
- **THEN** it returns concatenated complete `w:p` XML strings without visible-text projection, reserved-marker escaping, declaration, or wrapper

### Requirement: Text discovery APIs
The system SHALL expose `find_text(pattern, regex=False, occurrence=None, paragraph_id=None, raw=False)` and `count_matches(...)` with identical literal and regex semantics. A `TextMatch` MUST report paragraph ID, character range, XML-node range, whether the match crosses nodes, and total match count. Raw-mode ranges and truncation MUST be measured against XML strings.

#### Scenario: Cross-node text discovery
- **WHEN** a visible pattern spans two text nodes
- **THEN** `find_text` reports one match with both node range and `crosses_nodes=true`, and `count_matches` reports the same total

### Requirement: Raw query separation
All paragraph content read APIs SHALL accept `raw`, defaulting to `false`; result items MUST expose exactly one of `text` and `xml`. Raw data MUST be treated as trusted-caller output and MUST NOT enter the LLM context. APIs that neither read nor write content MUST NOT accept a meaningless `raw` argument.

#### Scenario: Listing field exclusivity
- **WHEN** the same paragraph is listed once in visible mode and once in raw mode
- **THEN** the visible item contains only `text`, the raw item contains only `xml`, and the raw value includes the `w:p` tags

### Requirement: Table position metadata
Paragraph metadata for table cells SHALL include global table index, row index, physical cell index, logical column index, grid span, nesting depth, and paragraph index in the cell. Logical columns MUST account for `w:gridSpan` and `w:gridBefore`; nested tables MUST have correct global indexes and depths. Layout that cannot be reliably reconstructed, including vertical merges, MUST be represented explicitly and MUST NOT be guessed. Location metadata MUST never participate in targeting.

#### Scenario: Spanned cell advances logical column
- **WHEN** a row's preceding cell has `gridSpan=2`
- **THEN** the following cell's logical column starts after both spanned grid columns while its paragraph remains targeted only by ID

