## 1. Package foundation and contracts

- [x] 1.1 Inspect the reference implementation and repository constraints, then choose and document the supported Python versions, build backend, package layout, and mature dependency set.
- [x] 1.2 Create the installable package, test layout, formatter, linter, type checker, pytest configuration, 60-second default timeout, and branch-coverage enforcement at 85%.
- [x] 1.3 Define the minimal public exports, immutable query/location models, edit/result models, selectors, content sources, save result, and JSON schema validation boundaries.
- [x] 1.4 Implement the complete `DocxKnifeError` hierarchy with bounded messages, serializable fields, cause preservation, and contract tests for every public exception.

## 2. Document loading, discovery, and anchors

- [x] 2.1 Implement `Document.open`, context-manager cleanup, DOCX/ZIP validation, private temporary workspace management, and source fingerprint capture.
- [x] 2.2 Parse only `word/document.xml` with secure XML settings and build document-order traversal for body and table-cell paragraphs while excluding SDT descendants from editing.
- [x] 2.3 Implement the anchor manifest with deterministic instance-local IDs, live-node validation, monotonic allocation, deletion/invalidation, and no fuzzy relocation or ID reuse.
- [x] 2.4 Build paragraph style and location metadata, including SDT structural occupancy, nested-table global indexes, `gridBefore`, `gridSpan`, logical columns, vertical-merge metadata, and nesting depth.
- [x] 2.5 Implement `paragraph_count`, paginated `list_paragraphs`, `get_paragraph`, and `get_visible_text` with visible/raw field exclusivity and exact truncation semantics.
- [x] 2.6 Implement literal/regex `grep_paragraphs`, `find_text`, and `count_matches`, including complete-value search before preview truncation and structured match ranges.
- [x] 2.7 Add focused unit and real-DOCX tests for missing/duplicate `w14:paraId`, SDT exclusion, stale anchors, pagination, raw queries, nested tables, spans, and location metadata.

## 3. TextMap and paragraph-internal editing

- [ ] 3.1 Implement final-view visible-text extraction across `w:t` and runs, including `w:ins`, excluding `w:del`, and mapping each character to node, offset, global offset, and run.
- [ ] 3.2 Implement reversible projection for tab, line/page/column breaks and carriage returns, literal `[[DOCX:` escaping, atomic ranges, and invalid-marker rejection.
- [ ] 3.3 Record intersected fields, hyperlinks, bookmarks, revisions, and other protected structures in match metadata and codify the allow/preserve/reject capability matrix with real OOXML fixtures.
- [ ] 3.4 Implement literal and regex selector compilation, invalid-pattern errors, zero-length-match policy, unique-match enforcement, indexed occurrence, and `occurrence=-1` right-to-left execution.
- [ ] 3.5 Implement `replace_text`, `delete_text`, `insert_text_before`, and `insert_text_after` across nodes/runs, rebuilding TextMap after each operation.
- [ ] 3.6 Preserve unmatched nodes and formatting, inherit insertion/replacement run properties, restore reserved nodes, and remove only text/run nodes that are structurally safe to remove.
- [ ] 3.7 Reject raw paragraph-internal edits and protected/atomic boundary crossings without destructive reconstruction.
- [ ] 3.8 Add unit and property tests for randomized run splitting, cross-run regex/literal matches, every occurrence mode, marker round trips, formatting inheritance, sequential edits, and protected structures.

## 4. Paragraph operations and fluent API

- [ ] 4.1 Implement visible-mode item expansion, CRLF/CR normalization, single-break `w:br`, multi-break paragraph boundaries, `xml:space`, and anchor `w:pPr`/first ordinary `w:rPr` inheritance.
- [ ] 4.2 Implement ordered `insert_para_before` and moving-cursor `insert_para_after`, assigning manifest IDs in final document order.
- [ ] 4.3 Implement one-to-many `replace_para` with old-ID invalidation, detected-structure warnings, and ordered replacement IDs.
- [ ] 4.4 Implement duplicate-free multi-target `delete_para` with complete prevalidation, reverse-document-order removal, and manifest synchronization.
- [ ] 4.5 Implement raw fragment parsing through a namespace-safe wrapper, require one or more top-level WordprocessingML `w:p` nodes, preserve supplied internals/order, and reject mixed modes.
- [ ] 4.6 Implement document-level paragraph methods and fluent `Paragraph` methods, returning live `Paragraph` objects that can be chained immediately.
- [ ] 4.7 Add real-DOCX tests for before/after order, item expansion, replacement/deletion, formatting inheritance, warning capture, raw round trips, invalid fragments, tables, and chained anchors.

## 5. Content resolution and normalization

- [x] 5.1 Implement exclusive `content_literal`/`content_ref` validation and a resolver interface that completes before any DOM mutation.
- [x] 5.2 Implement single-value JSONPath resolution with explicit source loading and structured missing/multi-value failures.
- [x] 5.3 Implement file references with configured allowed input roots, encoding validation, path-containment enforcement, and traversal/symlink security tests.
- [x] 5.4 Implement command references with argv-only execution, no shell, workspace-confined working directory, bounded stdout, controlled environment policy, timeout termination, exit-code checks, and UTF-8 validation.
- [x] 5.5 Apply identical visible-mode newline expansion to literal, JSONPath, file, and command results while treating raw-mode newlines only as XML whitespace.
- [x] 5.6 Specify and implement deterministic `normalize_text=true` punctuation/spacing rules with URL, email, and code-span protection; preserve exact input by default and never trim boundary spaces.
- [x] 5.7 Add failure and edge tests for source cardinality, path escape, decoding, JSONPath shapes, command timeout/nonzero/invalid UTF-8/output limits, newline runs, and protected-token normalization.

## 6. Batch transaction, validation, and audit

- [ ] 6.1 Implement whole-batch schema/ID/occurrence/content/raw prevalidation and deterministic normalization before mutation.
- [ ] 6.2 Implement the complete target-set conflict matrix, same-direction insertion merging, and dependency ordering for insertion, replacement, deletion, and paragraph-internal operations.
- [ ] 6.3 Implement transaction snapshots for DOM, anchor manifest, ID allocator, and change log with exact restoration on execution or precommit failure.
- [ ] 6.4 Implement ordered batch execution and success-only `EditResult`/`OperationResult`, including new IDs, warnings, bounded previews, and input-order result correlation.
- [ ] 6.5 Implement precommit checks for consumed operations, target outcomes, paragraph-count deltas, reparsable XML, logged warnings/results, and canonical equivalence of untouched structures.
- [ ] 6.6 Implement bounded success audit events and a single failed-batch event with expected/actual state, original cause, and `rolled_back=true`.
- [ ] 6.7 Add conflict-matrix, full-prevalidation, mid-batch failure, validation failure, state-restoration, audit-boundary, and canonical-fidelity tests, including property tests for failed-batch invariance.

## 7. Safe package persistence

- [ ] 7.1 Implement save-time source fingerprint comparison, including explicit behavior when source and output paths are identical and across repeated saves.
- [ ] 7.2 Implement atomic replacement of `<output>.bak` from the current destination before overwrite, with nullable backup results for new destinations.
- [ ] 7.3 Serialize `word/document.xml` without pretty printing and rebuild a temporary DOCX while preserving all other ZIP entry names, metadata, order where supported, and uncompressed content.
- [ ] 7.4 Reopen the temporary ZIP, parse the main XML, run package checks, and atomically replace the destination while preserving the original destination on failure.
- [ ] 7.5 Add integration tests for source drift, same-path save, new and repeated destination saves, backup replacement, injected write/rename failures, ZIP/XML revalidation, and unchanged non-main entries.

## 8. Public delivery and quality gates

- [ ] 8.1 Build end-to-end real-DOCX scenarios covering open, query, ID-targeted mixed edits, rollback, save, backup, reopen, and semantic/structural fidelity.
- [ ] 8.2 Add performance benchmarks for large-document pagination, paragraph/table indexing, TextMap construction, snapshot cost, and batch editing; eliminate repeated whole-DOM scans.
- [ ] 8.3 Configure CI for formatting, linting, type checking, unit/integration/property tests, branch coverage, and the supported Python-version matrix.
- [ ] 8.4 Write API reference and quickstart documentation for every public type, operation, raw trusted-caller boundary, error, backup behavior, and lifecycle guarantee.
- [ ] 8.5 Create the Agent Skill using only public APIs and a `raw=false` schema, enforcing returned IDs and preferring `content_ref` for long or deterministic data.
- [ ] 8.6 Run all quality gates, build/install the distribution in a clean environment, execute the quickstart and Agent Skill smoke tests, and record benchmark baselines.
