# Lifecycle

Every use of `docx-knife` walks the same four-stage lifecycle:

```
Document.open  →  query APIs  →  batch_edit  →  save
```

Each stage has hard guarantees. Do not skip stages, do not interleave
them, and do not reuse IDs across `open`/`close` boundaries.

## 1. `Document.open`

[`Document.open`](api.md#docx_knife.Document) validates the source path, opens the
DOCX as a ZIP, extracts every part into a **private per-instance
temporary directory**, parses `word/document.xml` with a secure XML
parser (external entities disabled, DTDs off), and captures a source
fingerprint:

- SHA-256 of the source bytes
- file size
- `mtime_ns`

The fingerprint is used later by `save` to detect out-of-band changes.

The returned object is a context manager. The private workspace is
released by `close()` (idempotent) or automatically when the `with`
block exits:

```python
with Document.open("in.docx") as doc:
    ...
```

If the path does not exist, `Document.open` raises
[`DocumentNotFoundError`](errors.md). If the file is not a valid DOCX
ZIP or `word/document.xml` cannot be parsed, it raises
[`InvalidDocumentError`](errors.md).

## 2. Query stage

Read APIs are the only sanctioned source of paragraph IDs:

- `Document.paragraph_count`
- `Document.list_paragraphs` — paginated; previews truncated at `max_chars`
- `Document.get_paragraph`
- `Document.get_visible_text`
- `Document.grep_paragraphs`
- `Document.count_matches`
- `Document.find_text`
- `Document.get_paragraph_object`

IDs are instance-local, allocated monotonically, and stable while the
`Document` is alive. Deleted or replaced paragraphs invalidate their IDs
permanently; IDs are never reused. IDs are meaningless across processes,
files, or reopens — a new `Document.open` re-numbers everything.

Paragraphs living inside `<w:sdt>` structured document tags are visible
to read APIs (as regular paragraphs from the caller's perspective) but
some structural editing may be constrained by the executor to preserve
the tag's contract.

## 3. `batch_edit`

`Document.batch_edit` executes an ordered sequence of operations under a
single transaction:

1. **Prevalidation.** JSON schema (if the batch came from JSON via
   `validate_batch`), ID validity, selector compile, `occurrence`
   bounds, content-source cardinality, `raw` field acceptance, and the
   target-set conflict matrix are verified **before** any DOM mutation.
2. **Content resolution.** Every `content_ref` is loaded (JSONPath, file,
   command) to a resolved string, up front. Later DOM edits cannot
   trigger new I/O.
3. **Snapshot.** DOM, anchor manifest, ID allocator, change-log, and
   warning state are captured.
4. **Apply.** Operations execute in ordered form. Successful operations
   yield `OperationResult` records correlated by input order.
5. **Precommit checks.** Consumed operations, target outcomes,
   paragraph-count delta, reparsable XML, and canonical equivalence of
   untouched structures are asserted.
6. **Commit or rollback.** On success the snapshot is discarded and an
   `EditResult` is returned. On any failure the snapshot is restored
   exactly and a [`BatchOperationError`](errors.md) with
   `rolled_back=True` is raised. There is no partial success.

Batches are the only supported mutation entry point. Do not compose
several `batch_edit` calls to emulate one big change unless you
intentionally want each subset to be independently atomic. Recommended
upper bound: 50 operations per batch.

## 4. `save`

`Document.save` persists the in-memory DOM back to a `.docx` file with
these guarantees:

1. **Source drift check.** The current source file is re-fingerprinted
   and compared byte-for-byte against the fingerprint captured at
   `open`. Any change raises [`SourceChangedError`](errors.md) and
   nothing is written.
2. **Rebuild.** `word/document.xml` is serialized without pretty-print;
   every other ZIP entry (media, styles, headers, `[Content_Types].xml`,
   relationships…) is copied byte-for-byte with entry order and
   metadata (timestamp, external attrs, comment, extra) preserved.
3. **Revalidate.** The rebuilt package is reopened; `ZipFile.testzip()`
   plus a secure XML re-parse of the main part must succeed.
4. **Backup.** If the destination already exists, its current contents
   are copied to `<output>.bak` **atomically** (via a temp file plus
   `os.replace`). A new destination produces `backup_path=None`.
5. **Atomic replace.** The rebuilt package is `os.replace`-d into the
   destination. Any failure before this step leaves the original
   destination untouched.
6. **Same-path re-save.** When output equals source, the source
   fingerprint is refreshed from disk so subsequent saves in the same
   session continue to succeed.

Every successful save returns a `SaveResult` with the resolved output
path, the backup path (or `None`), and any structural warnings recorded
during the run.

## Guarantees summary

| Guarantee | Enforced by |
| --- | --- |
| No partial batch application | `batch_edit` snapshot + rollback |
| Rebuild does not mutate untouched ZIP entries | `save` copies raw bytes and clones `ZipInfo` |
| No overwrite without an atomic `.bak` | `save` writes `.bak` via temp+rename before replacing destination |
| No silent overwrite when source changed on disk | `save` re-fingerprints and raises `SourceChangedError` |
| IDs never point to stale nodes | Anchor manifest invalidates on delete/replace; live-node resolve on every operation |
| No orphaned temp files | `Document.close` (or context-manager exit) removes the private workspace |
