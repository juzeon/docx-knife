## ADDED Requirements

### Requirement: Minimal document lifecycle
The public lifecycle SHALL consist of `Document.open(source_path)` and `Document.save(output_path)`, with context-manager support. `open` MUST reject missing or invalid DOCX inputs through structured errors and record a source fingerprint while parsing `word/document.xml` in a private temporary workspace.

#### Scenario: Invalid source package
- **WHEN** the source is missing or its main document XML cannot be parsed
- **THEN** `DocumentNotFoundError` or `InvalidDocumentError` is raised with the source path and reason

### Requirement: Source drift protection
Before every save, the system SHALL compare the source against its open-time fingerprint. If the source changed externally, it MUST raise `SourceChangedError` and MUST NOT overwrite the source, target, or external version.

#### Scenario: Source changes after open
- **WHEN** another process modifies the source DOCX before save
- **THEN** save raises `SourceChangedError` and leaves the externally modified file and destination unchanged

### Requirement: Backup before overwrite
If the output path exists, save SHALL atomically replace `<output_path>.bak` with the prior output before replacing the output. If output does not exist, no backup SHALL be created and `backup_path` SHALL be null. Repeated saves MUST retain only the immediately previous output version in `.bak`.

#### Scenario: Consecutive overwrite saves
- **WHEN** the same output is saved over twice
- **THEN** the final output is newest and its `.bak` is exactly the version replaced by the second save

### Requirement: Main-part-only package update
Save SHALL serialize `word/document.xml` without pretty printing, copy every other ZIP entry unchanged, replace only the main document entry, and MUST NOT add parts or modify content types or relationships.

#### Scenario: Package entries remain unchanged
- **WHEN** an edited document is saved
- **THEN** every ZIP entry other than `word/document.xml` has identical content and no entry is added or removed

### Requirement: Temporary output validation
Before destination replacement, save SHALL reopen the temporary ZIP and parse its `word/document.xml`. A ZIP or XML validation failure MUST raise `ValidationError(stage, checks, failed_check)` and MUST leave the original output intact.

#### Scenario: Serialized main XML is invalid
- **WHEN** the temporary package contains main XML that cannot be reparsed
- **THEN** save raises `ValidationError` and does not rename the temporary file over the output

### Requirement: Atomic destination replacement
After backup and temporary-package validation, save SHALL atomically rename the temporary output to its destination. If a later save stage fails, the original destination MUST remain available and any already completed backup MUST be preserved.

#### Scenario: Final replacement fails
- **WHEN** destination replacement fails after backup completion
- **THEN** the destination is not partially written and the completed `.bak` remains available for recovery

### Requirement: Save result contract
A successful save SHALL return `SaveResult` containing `output_path`, nullable `backup_path`, and warnings. Temporary paths, unpacking details, source fingerprints, and recovery implementation MUST remain private and MUST NOT appear in the Agent schema or change log.

#### Scenario: Save to new destination
- **WHEN** a document saves successfully to a nonexistent path
- **THEN** the result reports that output path, `backup_path=null`, and no private implementation paths

### Requirement: Untouched XML fidelity
For XML nodes outside operation targets, the system MUST preserve semantic content and structure. Fidelity comparison SHALL ignore namespace-prefix choice, attribute order, and equivalent parsed serialization, and MUST NOT claim byte-identical XML after parse and serialization.

#### Scenario: Equivalent serialization differs
- **WHEN** untouched XML is serialized with different prefixes or attribute order but parses to equivalent structure
- **THEN** fidelity validation accepts it while still rejecting semantic or structural changes

### Requirement: Unsupported document structures remain untouched
The engine SHALL neither edit nor synthesize SDTs, revision markup, headers, footers, footnotes, endnotes, comments, text boxes, images, or field refreshes. It MUST reject an edit whose scope traverses unsupported structure rather than destructively rebuilding it.

#### Scenario: Edit traverses unsupported content
- **WHEN** an requested edit would cross an unsupported structure
- **THEN** the operation raises `UnsupportedStructureError`, save is not required to repair anything, and the structure remains unchanged

