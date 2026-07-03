---
name: docx-knife
description: Edit .docx files without hallucinating text, breaking OOXML, or dumping the whole document into context. Use for any request that mutates paragraphs or paragraph-internal spans in a DOCX file.
---

# docx-knife

## 1. Purpose

`docx-knife` lets an LLM edit `.docx` files without hallucinating source text, breaking OOXML structure, or streaming the whole document into context. The skill enforces a **query-then-patch** protocol: first call read APIs to obtain stable per-document paragraph IDs (`p_000001`, `p_000042`, ...) and short previews, then submit a batch of structured operations that reference those IDs. The executor owns XML parsing, TextMap alignment, structural preservation, atomic rollback, and safe save with backup.

## 2. Prohibitions

- Never emit XML, XPath, array indexes, `w14:paraId`, character offsets, or any invented ID. Only IDs returned by the read APIs in the current session are valid.
- Never quote long source text back into `content_literal`. Anything longer than ~40 characters, and any number, date, monetary amount, company name, URL, or email address, must use `content_ref` (jsonpath / file / command).
- Never use `raw=true`. The LLM-facing schema forbids the `raw` field; raw XML is a trusted-caller-only channel.
- Never call `save()` in the middle of a batch. One `batch_edit(...)` call is one atomic write; call `save(output_path)` once at the end.
- Never retry a failed batch verbatim. Read the structured error, rebuild the batch (fresh IDs, corrected selectors), then submit again.

## 3. Workflow

1. **Locate.** Call `list_paragraphs`, `grep_paragraphs`, or `find_text` to obtain paragraph IDs and previews. Keep the IDs in scratch memory; never invent them.
2. **Compose.** Build a batch envelope of operations that reference those IDs. Prefer `content_ref` for long or deterministic text.
3. **Submit.** Call `batch_edit(operations=[...])`. On success you receive an `EditResult` whose `results` correlate with input order. On failure you receive a structured `DocxKnifeError` — inspect its fields, do not retry blindly.
4. **Save.** Call `save(output_path)` once. It writes atomically and produces a `.bak` of any pre-existing destination.

Each batch touches only the current `Document` instance's DOM. If the process reopens the file, all previously issued IDs are invalidated and must be re-queried.

## 4. Public API summary

Full reference: `docs/api.md`. One-line signatures the agent may call:

Read side (all return immutable models, no XML strings unless `raw=true`, which the LLM must not use):

- `Document.open(source_path, *, content_config=None) -> Document`
- `Document.close() -> None`  /  `Document` supports `with ...:` context management.
- `Document.paragraph_count() -> int`
- `Document.list_paragraphs(start=1, limit=None, max_chars=80, raw=False) -> ParagraphListResult`
- `Document.get_paragraph(paragraph_id, raw=False) -> str`
- `Document.get_visible_text(raw=False) -> str`
- `Document.grep_paragraphs(pattern, regex=False, start=1, limit=None, max_chars=0, raw=False) -> ParagraphSearchResult`
- `Document.count_matches(pattern, regex=False, paragraph_id=None, raw=False) -> int`
- `Document.find_text(pattern, regex=False, occurrence=None, paragraph_id=None, raw=False) -> TextMatch | list[TextMatch] | None`
- `Document.get_paragraph_object(paragraph_id) -> Paragraph`

Write side (one call per batch — do not interleave writes across batches):

- `Document.batch_edit(operations, *, normalize_text=False, envelope=None) -> EditResult`
- `Document.save(output_path) -> SaveResult`
- `Document.change_log() -> list[dict]` — audit only, not for control flow.

Schema helpers exported from the top-level package:

- `docx_knife.BATCH_SCHEMA` — the JSON Schema below.
- `docx_knife.validate_batch(payload) -> None` — raises `ValidationError` on failure.

## 5. JSON schema for `batch_edit`

Every operation carries a unique `op_id` (string), an `op` discriminator, one or more target IDs, and either `items` (paragraph ops) or `find` plus optional content (text ops). `additionalProperties` is `false` at every level; the `raw` field is not permitted. Machine-readable copy: `agent_schema.json` in this directory.

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "docx-knife batch",
  "type": "object",
  "additionalProperties": false,
  "required": ["operations"],
  "properties": {
    "operations": {
      "type": "array",
      "minItems": 1,
      "items": {
        "oneOf": [
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "items"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "insert_para_before"},
              "target_id": {"type": "string", "minLength": 1},
              "items": {"$ref": "#/$defs/items"}
            }
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "items"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "insert_para_after"},
              "target_id": {"type": "string", "minLength": 1},
              "items": {"$ref": "#/$defs/items"}
            }
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "items"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "replace_para"},
              "target_id": {"type": "string", "minLength": 1},
              "items": {"$ref": "#/$defs/items"}
            }
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_ids"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "delete_para"},
              "target_ids": {
                "type": "array",
                "minItems": 1,
                "uniqueItems": true,
                "items": {"type": "string", "minLength": 1}
              }
            }
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "find"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "replace_text"},
              "target_id": {"type": "string", "minLength": 1},
              "find": {"$ref": "#/$defs/find"},
              "occurrence": {"type": "integer", "minimum": -1},
              "content_literal": {"type": "string"},
              "content_ref": {"$ref": "#/$defs/content_ref"}
            },
            "oneOf": [
              {"required": ["content_literal"], "not": {"required": ["content_ref"]}},
              {"required": ["content_ref"], "not": {"required": ["content_literal"]}}
            ]
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "find"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "delete_text"},
              "target_id": {"type": "string", "minLength": 1},
              "find": {"$ref": "#/$defs/find"},
              "occurrence": {"type": "integer", "minimum": -1}
            }
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "find"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "insert_text_before"},
              "target_id": {"type": "string", "minLength": 1},
              "find": {"$ref": "#/$defs/find"},
              "occurrence": {"type": "integer", "minimum": -1},
              "content_literal": {"type": "string"},
              "content_ref": {"$ref": "#/$defs/content_ref"}
            },
            "oneOf": [
              {"required": ["content_literal"], "not": {"required": ["content_ref"]}},
              {"required": ["content_ref"], "not": {"required": ["content_literal"]}}
            ]
          },
          {
            "type": "object",
            "additionalProperties": false,
            "required": ["op_id", "op", "target_id", "find"],
            "properties": {
              "op_id": {"type": "string", "minLength": 1},
              "op": {"const": "insert_text_after"},
              "target_id": {"type": "string", "minLength": 1},
              "find": {"$ref": "#/$defs/find"},
              "occurrence": {"type": "integer", "minimum": -1},
              "content_literal": {"type": "string"},
              "content_ref": {"$ref": "#/$defs/content_ref"}
            },
            "oneOf": [
              {"required": ["content_literal"], "not": {"required": ["content_ref"]}},
              {"required": ["content_ref"], "not": {"required": ["content_literal"]}}
            ]
          }
        ]
      }
    }
  }
}
```

The `agent_schema.json` file in this directory holds the fully expanded schema exactly as `docx_knife._schema.BATCH_SCHEMA` produces it; a smoke test guards against drift.

## 6. `content_ref` forms

Use `content_ref` (never `content_literal`) for any deterministic value: prices, dates, party names, statistics, template clauses, generated output. Exactly one form per item:

**jsonpath** — extract a single value from a JSON file. Multi-value hits are rejected.
```json
{"content_ref": {"type": "jsonpath", "source": "contract.json", "path": "$.party_a.name"}}
```

**file** — read a UTF-8 (or specified-encoding) text file whose path resolves inside the configured input roots.
```json
{"content_ref": {"type": "file", "path": "clauses/confidentiality.txt", "encoding": "utf-8"}}
```

**command** — capture stdout of a subprocess launched with argv only.
```json
{
  "content_ref": {
    "type": "command",
    "argv": ["python", "scripts/render_clause.py", "--contract", "contract.json"],
    "timeout_seconds": 30
  }
}
```

Constraints, enforced by the resolver:

- `file.path` must resolve inside an allowed input root; symlink escape and path traversal are rejected.
- `command.argv` runs without a shell; the cwd is workspace-confined; `timeout_seconds` is mandatory (must be `> 0`); non-zero exit, non-UTF-8 output, or exceeding the output limit is rejected.
- `jsonpath` must resolve to a single value; missing keys and multi-value hits are rejected.

Newline handling for visible-mode content: `\r\n`/`\r` are normalized to `\n`; a single `\n` becomes `<w:br/>`; two or more consecutive `\n` split into paragraph boundaries. This applies identically to `content_literal`, jsonpath, file, and command results, so one item can expand into several paragraphs.

## 7. Error handling

Every failure raises a subclass of `DocxKnifeError` with structured fields. Recommended agent responses:

| Error | Cause | Recommended response |
| --- | --- | --- |
| `ParagraphNotFoundError(target_id)` | The referenced ID no longer exists in the DOM. | Refetch IDs via `list_paragraphs`/`grep_paragraphs` and rebuild the batch. Do not reuse invalidated IDs. |
| `AmbiguousTextMatchError(target_id, selector, total_matches)` | `find` matched more than one span and `occurrence` was omitted. | Either supply an explicit `occurrence` (0-based, or `-1` for all) or narrow the selector so it is unique. |
| `TextNotFoundError(target_id, selector, occurrence, total_matches)` | The selector matched fewer times than `occurrence` requires. | Verify the text still exists via `find_text`/`count_matches`. Do not retry blindly with different text. |
| `BatchOperationError(operation_index, op_id, reason, cause, rolled_back=true)` | An operation failed and the whole batch was rolled back. | Inspect `.reason` and `.cause`; do not resubmit the same batch. Fix the offending op and rebuild. |
| `SourceChangedError(source_path)` | The source file changed on disk since `Document.open`. | Close and reopen the document, re-query IDs, and rebuild the batch. |
| `InvalidContentError(raw, reason)` | A `content_literal`/`content_ref` failed validation (missing key, bad encoding, unknown `[[DOCX:...]]` marker, etc.). | Fix the referenced content or the literal; do not retry unchanged. |
| `UnsupportedStructureError(target_id, structures, matched_range)` | The edit range crosses a protected structure (field, hyperlink, bookmark, revision, ...) that the executor refuses to break. | Narrow the edit scope or target a different paragraph/selector. Do not force through. |
| `InvalidPatternError(pattern, reason)` | A regex selector failed to compile. | Fix the pattern (or set `regex=false`) and re-submit. |
| `ValidationError(stage, checks, failed_check)` | Batch schema, precommit, or post-commit validation failed. | Do not retry. Escalate; the input violates a structural invariant. |
| `DocumentNotFoundError(path)` / `InvalidDocumentError(path, reason)` | The DOCX cannot be opened. | Fix the path or file; nothing to retry. |

`BatchOperationError.rolled_back` is always `true`: the document state is exactly as it was before the batch. Never assume partial success.

## 8. Example: end-to-end batch

Replace one dated span, insert a two-paragraph clause after the anchor, and drop a stale paragraph — all in one atomic batch after IDs have been discovered via `grep_paragraphs`:

```json
{
  "operations": [
    {
      "op_id": "op_001",
      "op": "replace_text",
      "target_id": "p_000042",
      "find": {"pattern": "30 天", "regex": false},
      "occurrence": 0,
      "content_literal": "60 天"
    },
    {
      "op_id": "op_002",
      "op": "insert_para_after",
      "target_id": "p_000042",
      "items": [
        {"content_ref": {"type": "file", "path": "clauses/confidentiality_v2.txt", "encoding": "utf-8"}},
        {"content_literal": "本条款自签署之日起生效。"}
      ]
    },
    {
      "op_id": "op_003",
      "op": "delete_para",
      "target_ids": ["p_000051"]
    }
  ]
}
```

## 9. Constraints for LLM output

- Every operation has a unique `op_id` inside the batch.
- Every `target_id` / `target_ids` value comes from a read call issued in **this** session.
- Prefer `content_ref` when the value is longer than ~40 characters, or when it is a number, date, monetary amount, party name, URL, or email — anything the model should not paraphrase.
- Keep batches to ≤ 50 operations. Split larger changes into successive batches (each atomic) rather than one giant envelope.
- Never set `raw`, and never emit XML fragments; the schema will reject them.
- After a `batch_edit` succeeds, call `save(output_path)` exactly once to persist the change (with automatic `.bak` of any prior destination file).
