# text-map-editing Specification

## Purpose

Define paragraph-internal editing on top of a reversible visible-text map: continuous cross-run projection, reserved markers for non-text nodes, final-view revision handling, selector semantics, deterministic match disambiguation, format-preserving writeback, and rejection of edits that cross protected structures or use raw mode.

## Requirements

### Requirement: Reversible visible-text map
For each editable paragraph, the system SHALL build a `TextMap` whose continuous `text` maps every character to XML node, node offset, global text offset, and owning run. The map MUST span adjacent `w:t` elements and runs and MUST expose atomic ranges.

#### Scenario: Text split across runs
- **WHEN** `违约` and `责任` occur in separate runs
- **THEN** the map exposes `违约责任` as continuous text and maps every character back to its original node and run

### Requirement: Non-text visible-node projection
The map SHALL project tab, ordinary line break, page break, column break, and carriage return as `[[DOCX:TAB]]`, `[[DOCX:LINE_BREAK]]`, `[[DOCX:PAGE_BREAK]]`, `[[DOCX:COLUMN_BREAK]]`, and `[[DOCX:CR]]`. Literal source text beginning `[[DOCX:` MUST be escaped as `\[[DOCX:` and round-trip without loss. Each reserved marker MUST be indivisible; a match boundary inside one MUST cause the operation to fail.

#### Scenario: Reserved nodes round-trip
- **WHEN** a paragraph contains every supported non-text visible node and literal `[[DOCX:` text
- **THEN** projection and writeback restore each original node type and the literal text without ambiguity

### Requirement: Final-view extraction and structure metadata
Visible text SHALL use Word's final view by including `w:ins` content and excluding `w:del` content. Fields, hyperlinks, bookmarks, and related structures intersecting a match MUST be recorded in match metadata for capability-policy evaluation.

#### Scenario: Existing revision content
- **WHEN** a paragraph contains inserted and deleted revision text
- **THEN** searches see inserted text, do not see deleted text, and report intersected protected structures

### Requirement: Text selector semantics
Text operations SHALL use a selector containing `pattern` and `regex`, with a string accepted as literal shorthand. Literal mode MUST match exact text and regex mode MUST evaluate the regular expression across the continuous mapped text. Invalid regular expressions MUST raise `InvalidPatternError`.

#### Scenario: Non-greedy cross-run selector
- **WHEN** a non-greedy regex identifies a unique range whose endpoints lie in different runs
- **THEN** the system selects exactly that continuous visible-text range

### Requirement: Match disambiguation
After resolving `target_id`, the system SHALL calculate all selector matches. Omitted `occurrence` MUST require exactly one match and otherwise raise `AmbiguousTextMatchError`; a nonnegative occurrence MUST select the zero-based match or raise `TextNotFoundError` when out of range; `occurrence=-1` MUST select all matches. The system MUST NOT automatically choose similar text.

#### Scenario: Ambiguous omitted occurrence
- **WHEN** the target paragraph contains two exact matches and occurrence is omitted
- **THEN** the operation raises `AmbiguousTextMatchError` with total match count and changes nothing

#### Scenario: Replace all
- **WHEN** occurrence is `-1` for a selector with multiple matches
- **THEN** all matches are edited from right to left so earlier character positions remain valid

### Requirement: Paragraph-internal editing operations
The system SHALL provide `replace_text`, `delete_text`, `insert_text_before`, and `insert_text_after`, each bound to one paragraph ID. These operations MUST support matches crossing text nodes and runs, and MUST rebuild the paragraph `TextMap` after every operation so later operations observe current text.

#### Scenario: Sequential edits use current text
- **WHEN** one operation changes a paragraph and a later operation targets the newly produced text
- **THEN** the later operation resolves against a rebuilt map and edits the new text

### Requirement: Format-preserving text writeback
Inserted text SHALL inherit the insertion-point run properties, and replacement text SHALL inherit the first replaced character's run properties. Complete known reserved markers MUST be restored to corresponding XML nodes; incomplete or unknown `[[DOCX:...]]` markers MUST raise `InvalidContentError`. Empty `w:t` nodes MAY be removed, but a run that still contains drawing, tab, break, or other content MUST be retained.

#### Scenario: Unmatched formatting survives replacement
- **WHEN** a cross-run match is replaced
- **THEN** replacement inherits the first matched run properties and all unmatched runs, formatting, and non-text nodes remain unchanged

### Requirement: Protected-structure rejection
If a selected text range or edit boundary crosses an unsupported or atomic structure, the entire operation MUST fail with `UnsupportedStructureError` and MUST NOT fall back to paragraph reconstruction.

#### Scenario: Match crosses protected structure
- **WHEN** a selected range traverses a protected field or reserved marker boundary
- **THEN** the item fails with structure and range metadata and leaves the paragraph unchanged

### Requirement: Text operations exclude raw mode
Paragraph-internal operations SHALL operate only on visible text and MUST reject `raw=true`. XML-level modification MUST be performed by reading raw paragraph XML and replacing the paragraph through the raw paragraph API.

#### Scenario: Raw text edit is rejected
- **WHEN** a caller requests `replace_text` with `raw=true`
- **THEN** the system rejects the operation before DOM modification
