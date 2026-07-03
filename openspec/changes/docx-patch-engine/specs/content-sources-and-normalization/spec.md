## ADDED Requirements

### Requirement: Exclusive content source
Each JSON paragraph item SHALL contain exactly one of `content_literal` and `content_ref`. Missing sources or both sources together MUST be rejected during prevalidation. `content_literal` SHALL provide its value directly; paragraph object APIs SHALL accept equivalent direct strings.

#### Scenario: Invalid source cardinality
- **WHEN** an item provides neither source or provides both
- **THEN** the batch is rejected before resolving or writing any content

### Requirement: Controlled JSONPath source
A JSONPath content reference SHALL identify a declared source and path and MUST resolve to exactly one required value. Missing keys and multi-value results where one value is required MUST raise a structured content error.

#### Scenario: JSONPath resolves multiple values
- **WHEN** a single-value item uses a JSONPath that returns multiple values
- **THEN** content resolution fails before DOM modification

### Requirement: Controlled file source
A file content reference SHALL read the declared path using the declared encoding and MUST require the path to lie within an allowed input root. Missing files, disallowed paths, and decoding failures MUST be rejected.

#### Scenario: File escapes allowed root
- **WHEN** a file reference resolves outside every configured input root
- **THEN** the system rejects it without reading its content

### Requirement: Shell-free command source
A command reference SHALL execute an argument array directly without a shell, restrict its working directory to the workspace, enforce `timeout_seconds`, and require successful exit and UTF-8 standard output. Timeout, nonzero exit, invalid UTF-8, or malformed argv MUST reject the operation.

#### Scenario: Command times out
- **WHEN** a referenced command exceeds its timeout
- **THEN** the command is terminated and content resolution fails without editing the document

#### Scenario: Shell syntax is not interpreted
- **WHEN** argv contains shell metacharacters
- **THEN** they are passed as literal arguments and no shell expansion or chaining occurs

### Requirement: Source resolution precedes mode processing
The executor SHALL resolve all references before applying the operation's raw-mode rules. Visible-mode resolved values MUST follow text newline and normalization rules; raw-mode values MUST be parsed as XML fragments. Raw mode MUST support literal, file, and command sources but MUST NOT expose a template registry or template reference type.

#### Scenario: Referenced raw fragment
- **WHEN** a raw file reference resolves to multiple complete `w:p` elements
- **THEN** the system validates and inserts them as XML rather than converting their newlines to breaks or paragraph boundaries

### Requirement: Deterministic newline expansion
In visible mode, the system SHALL normalize CRLF and CR to LF, convert a single LF within a paragraph to `w:br`, and treat each run of two or more LFs as one paragraph boundary. This behavior MUST be identical for literal, JSONPath, file, and command content and MAY cause one item to expand into multiple paragraphs. Raw mode MUST treat newlines only as XML whitespace.

#### Scenario: Mixed newline forms
- **WHEN** visible input contains CRLF, a single line break, and a run of multiple line breaks
- **THEN** all newline forms are normalized, the single break becomes `w:br`, and the run creates exactly one paragraph boundary

### Requirement: Optional text normalization
`normalize_text=false` SHALL be the default and MUST preserve punctuation, spaces, quotation marks, mixed-language text, and leading and trailing spaces except for required XML escaping and space preservation. When `normalize_text=true`, the system SHALL apply deterministic basic Chinese punctuation and Chinese/Latin spacing normalization while skipping URLs, email addresses, and code spans, and MUST NOT remove leading or trailing spaces.

#### Scenario: Default leaves prose unchanged
- **WHEN** visible mixed Chinese and Latin text is written without enabling normalization
- **THEN** its punctuation and spacing remain unchanged apart from OOXML-required representation

#### Scenario: Protected token normalization
- **WHEN** normalization is enabled for text containing a URL, email, code span, and surrounding Chinese prose
- **THEN** eligible prose is normalized while the protected tokens and boundary spaces remain byte-for-byte text-equivalent

