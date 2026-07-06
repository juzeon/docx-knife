# atomic-batch-processing Specification

## Purpose

Define the batch execution boundary for docx-knife writes: a uniform operations envelope, deterministic normalization and prevalidation, conflict-safe merging and ordering, atomic apply-with-rollback semantics, structured errors, a bounded change log, commit-time validation, and a strict LLM-facing operation surface.

## Requirements

### Requirement: Uniform batch envelope and results
All writes SHALL use an `operations` batch envelope. Each operation MUST carry a stable `op_id`, operation kind, ID-based target fields, content mode where relevant, and operation-specific arguments. A successful `EditResult` SHALL represent the whole committed batch and contain one `OperationResult` per item, all with `status=success`; mixed success and failure results MUST NOT be returned.

#### Scenario: Fully successful batch
- **WHEN** every validated operation completes and commit validation passes
- **THEN** one result is returned for every operation in input order and every status is `success`

### Requirement: Input normalization and validation
Before mutation, the system SHALL normalize all operations and validate schemas, IDs, occurrences, nonempty items, content references, raw-mode consistency, and conflicts across the complete batch. Operations MUST apply only to the current in-memory `Document` and MUST NOT be reusable as serialized cross-process coordinates.

#### Scenario: Later operation is invalid
- **WHEN** the final operation in a batch has an invalid target
- **THEN** complete-batch prevalidation fails and none of the earlier valid operations execute

### Requirement: Deterministic conflict handling
The system SHALL merge repeated `insert_para_after` operations on one ID by operation order and SHALL do the same for repeated `insert_para_before`. It MUST reject multiple replacements of one ID, replacement plus deletion of one ID, deletion plus after-insertion on one ID, and any intersecting multi-target conflict implied by those rules. Deletion plus before-insertion SHALL execute insertion first; replacement plus either insertion direction SHALL execute insertion first and replacement second.

#### Scenario: Repeated after insertions merge
- **WHEN** multiple after-insertion operations target the same ID
- **THEN** their item arrays are merged in operation order and the resulting paragraphs retain that order

#### Scenario: Delete and after-insert conflict
- **WHEN** one batch deletes an ID and also uses it as an after-insertion anchor
- **THEN** prevalidation rejects the batch as conflicting

### Requirement: Atomic execution lifecycle
Batch execution SHALL normalize, fully prevalidate, snapshot DOM, manifest, and change log, apply operations in deterministic order, run precommit validation, and commit only if every stage succeeds. Any failure MUST restore all three states exactly to their pre-batch values.

#### Scenario: Mid-batch execution failure
- **WHEN** an operation fails after earlier operations mutated the in-memory document
- **THEN** DOM, manifest, and change log are restored to the batch-start state

### Requirement: Batch failure reporting
Any operation or validation failure SHALL raise `BatchOperationError` with operation index, `op_id`, human-readable reason, original cause, and `rolled_back=true`. A failed batch MUST record only one bounded audit event describing failure and rollback and MUST NOT return partial results.

#### Scenario: Cause is preserved
- **WHEN** content resolution raises a structured underlying exception during a batch
- **THEN** `BatchOperationError` identifies the operation, retains that cause, reports rollback, and no partial success result is returned

### Requirement: Structured public errors
Every public error SHALL inherit `DocxKnifeError` and provide serializable stable fields. The system MUST define `DocumentNotFoundError(path)`, `InvalidDocumentError(path, reason)`, `SourceChangedError(source_path)`, `ParagraphNotFoundError(target_id)`, `TextNotFoundError(target_id, selector, occurrence, total_matches)`, `AmbiguousTextMatchError(target_id, selector, total_matches)`, `InvalidPatternError(pattern, reason)`, `InvalidContentError(raw, reason)`, `UnsupportedStructureError(target_id, structures, matched_range)`, `BatchOperationError(...)`, and `ValidationError(stage, checks, failed_check)`. Human messages and candidate previews MUST be bounded; diagnostics MUST never trigger automatic write retries.

#### Scenario: Serializable ambiguous-match error
- **WHEN** a text operation is ambiguous
- **THEN** the raised error can be serialized with target ID, selector, and total matches while any preview is truncated

### Requirement: Bounded change log
Each successful operation SHALL log `op_id`, kind, targets, success status, warnings, and bounded before/after previews or counts. A failed batch SHALL log error type, expected and actual state, and `rolled_back=true`. The log MUST support review and regression diagnosis without storing unbounded full-document text.

#### Scenario: Multi-paragraph insertion log
- **WHEN** an insertion creates multiple paragraphs
- **THEN** its log records inserted count and bounded previews in document order without embedding the whole document

### Requirement: Commit validation
Before commit, the system SHALL verify that main XML is well formed and reparsable, every operation was consumed, target outcomes match expectations, paragraph-count deltas are correct, untouched XML is semantically and structurally equivalent under XML canonical comparison, and every result and warning is logged.

#### Scenario: Untouched structure changes unexpectedly
- **WHEN** an operation causes an unaddressed XML structure to differ beyond namespace prefixes, attribute order, or equivalent serialization
- **THEN** commit validation fails and the batch is fully rolled back

### Requirement: LLM operation boundary
The LLM-facing schema SHALL expose user intent, relevant visible paragraphs, available references, and only `raw=false` ID-targeted operations. It MUST require IDs returned by reads and MUST NOT expose or accept XML, XPath, indexes, `w14:paraId`, XML character offsets, targetless global old text, or invented IDs. The executor SHALL own disambiguation, mapping, structure protection, normalization, and rollback.

#### Scenario: LLM submits an index target
- **WHEN** an LLM-facing operation uses an array index instead of a returned paragraph ID
- **THEN** schema validation rejects it before execution
