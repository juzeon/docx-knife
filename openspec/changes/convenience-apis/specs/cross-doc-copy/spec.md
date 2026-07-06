## ADDED Requirements

### Requirement: Copy paragraphs from another document
The `Document` class SHALL provide a `copy_paragraphs_from(source, start_id, end_id)` method that extracts raw OOXML paragraph elements from a source `Document` instance (from `start_id` to `end_id` inclusive) and returns them as a `list[str]` of serialized `<w:p>` elements, suitable for use with `insert_para_before`/`insert_para_after` with `raw=True`.

#### Scenario: Copy a range of paragraphs
- **WHEN** `target.copy_paragraphs_from(source, "p_000010", "p_000015")` is called
- **THEN** a list of 6 serialized `<w:p>` XML strings is returned, each preserving the original run properties (bold, italic, font, etc.)

#### Scenario: Copy a single paragraph
- **WHEN** `target.copy_paragraphs_from(source, "p_000010", "p_000010")` is called
- **THEN** a list containing exactly 1 serialized `<w:p>` XML string is returned

#### Scenario: Invalid start or end ID
- **WHEN** `copy_paragraphs_from` is called with a `start_id` or `end_id` that does not exist in the source document
- **THEN** a `ParagraphNotFoundError` is raised referencing the invalid ID

#### Scenario: start_id appears after end_id
- **WHEN** `copy_paragraphs_from` is called with `start_id` positioned after `end_id` in document order
- **THEN** a `ValueError` is raised indicating invalid range order

### Requirement: Copied paragraphs include namespace declarations
Each serialized `<w:p>` string returned by `copy_paragraphs_from` SHALL include the `xmlns:w` namespace declaration so it can be directly used with `raw=True` insertion without additional namespace fixup.

#### Scenario: Namespace present in output
- **WHEN** paragraphs are copied from a source document
- **THEN** each returned string contains `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`

### Requirement: Run properties are preserved
Copied paragraphs SHALL retain all `<w:rPr>` elements (bold, italic, font name, size, color, underline, etc.) from the source paragraph's runs.

#### Scenario: Bold and italic runs preserved
- **WHEN** a source paragraph contains a bold run followed by an italic run
- **THEN** the serialized output preserves both `<w:b/>` and `<w:i/>` in their respective `<w:rPr>` elements
