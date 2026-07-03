# Lifecycle

```
Document.open  →  query APIs  →  batch_edit  →  save
```

Stages must not be interleaved. IDs are never reused across `open` / `close`.

## 1. `Document.open`

Validates the path, opens the DOCX ZIP, extracts every part into a **private per-instance temporary directory**, parses `word/document.xml` with a secure XML parser (external entities off, DTDs off), and captures a source fingerprint (SHA-256 + size + `mtime_ns`) used later by `save` to detect drift.

The result is a context manager; the workspace is released by `close()` (idempotent) or `with`-exit.

Errors: [`DocumentNotFoundError`](errors.md), [`InvalidDocumentError`](errors.md).

## 2. Query stage

Only read APIs mint valid IDs:

`paragraph_count`, `list_paragraphs`, `get_paragraph`, `get_visible_text`, `grep_paragraphs`, `count_matches`, `find_text`, `get_paragraph_object`.

IDs are instance-local, allocated monotonically, stable while the `Document` is alive, and permanently invalidated when their paragraph is deleted or replaced (never reused). A new `Document.open` re-numbers everything.

Paragraphs inside `<w:sdt>` structured document tags are visible to read APIs; some structural edits may be constrained to preserve the tag contract.

## 3. `batch_edit`

Ordered pipeline under a single transaction:

1. **Prevalidate.** JSON schema (if entered via `validate_batch`), ID validity, selector compile, `occurrence` bounds, content-source cardinality, `raw` acceptance, target-set conflict matrix.
2. **Resolve content.** All `content_ref` items are loaded up front (jsonpath / file / command). DOM mutations trigger no further I/O.
3. **Snapshot.** DOM, anchor manifest, ID allocator, change-log, warning state.
4. **Apply** ordered operations → per-op `OperationResult`.
5. **Precommit checks.** Consumed ops, target outcomes, paragraph-count delta, reparsable XML, canonical equivalence of untouched structures.
6. **Commit or rollback.** Success → `EditResult`; any failure → snapshot restored exactly and [`BatchOperationError(rolled_back=True)`](errors.md). No partial success.

Batches are the only mutation entry point. Recommended upper bound: 50 ops per batch.

## 4. `save`

1. **Drift check.** Re-fingerprint the source file; any change → [`SourceChangedError`](errors.md) and nothing is written.
2. **Rebuild.** Serialize `word/document.xml` without pretty-print; every other ZIP entry is copied byte-for-byte with entry order and metadata preserved.
3. **Revalidate.** Reopen the rebuilt package, `ZipFile.testzip()` + secure re-parse of the main part must succeed.
4. **Backup.** If the destination exists, its current contents are copied to `<output>.bak` atomically (temp file + `os.replace`); a new destination yields `backup_path=None`.
5. **Atomic replace.** `os.replace` the rebuilt package into the destination. Any earlier failure leaves the destination untouched.
6. **Same-path re-save.** Source fingerprint is refreshed so later saves in the session still succeed.

Returns `SaveResult(output_path, backup_path, warnings)`.

## Guarantees

| Guarantee | Enforced by |
| --- | --- |
| No partial batch application | snapshot + rollback in `batch_edit` |
| Untouched ZIP entries are byte-identical | `save` copies raw bytes and clones `ZipInfo` |
| No overwrite without atomic `.bak` | `save` writes `.bak` via temp+rename before replace |
| No silent overwrite on source drift | `save` re-fingerprints and raises `SourceChangedError` |
| IDs never point to stale nodes | anchor manifest invalidates on delete/replace; live resolve every op |
| No orphaned temp files | `close()` / context-manager exit removes the workspace |
