## ADDED Requirements

### Requirement: Document-wide text replacement
The `Document` class SHALL provide a `replace_all(find, replacement, *, regex=False, normalize_text=False)` method that replaces all occurrences of `find` across all paragraphs in the document and returns the total number of substitutions performed.

#### Scenario: Replace all occurrences of a literal string
- **WHEN** `doc.replace_all("Old Corp", "New Corp")` is called on a document with 5 paragraphs containing "Old Corp" (some paragraphs having multiple occurrences)
- **THEN** all occurrences across all paragraphs are replaced and the method returns the total count of replacements made

#### Scenario: Replace with regex pattern
- **WHEN** `doc.replace_all(r"第\d+条", "第X条", regex=True)` is called
- **THEN** all regex matches across all paragraphs are replaced and the total count is returned

#### Scenario: No matches found
- **WHEN** `doc.replace_all("nonexistent text", "replacement")` is called and no paragraph contains the search term
- **THEN** the method returns 0 and the document is unchanged

#### Scenario: Replacement preserves run formatting
- **WHEN** a replacement is performed on text that spans multiple runs with different formatting
- **THEN** the replacement text inherits the formatting of the leftmost matched run, consistent with `replace_text` behavior

### Requirement: Replace-all respects cross-run text
The `replace_all` method SHALL match text that spans across multiple `<w:r>` elements within a single paragraph, using the same TextMap logic as `replace_text`.

#### Scenario: Match spans two runs
- **WHEN** "hello world" is split across two runs ("hello " in run 1, "world" in run 2) and `replace_all("hello world", "hi earth")` is called
- **THEN** the match is found and replaced correctly
