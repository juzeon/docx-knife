"""Smoke tests for the docx-knife Agent Skill assets."""

from __future__ import annotations

import copy
import json
from pathlib import Path

import pytest

from docx_knife import BATCH_SCHEMA, ValidationError, validate_batch

SKILL_DIR = Path(__file__).resolve().parents[1] / "skills" / "docx-knife"


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def test_agent_schema_matches_source() -> None:
    exported = _load_json(SKILL_DIR / "agent_schema.json")
    assert exported == BATCH_SCHEMA, (
        "skills/docx-knife/agent_schema.json is out of sync with "
        "docx_knife._schema.BATCH_SCHEMA; regenerate via "
        "`python skills/docx-knife/_export_schema.py`."
    )


@pytest.mark.parametrize(
    "example",
    sorted((SKILL_DIR / "examples").glob("*.json")),
    ids=lambda p: p.name,
)
def test_examples_validate(example: Path) -> None:
    envelope = _load_json(example)
    validate_batch(envelope)


def test_schema_rejects_raw_field() -> None:
    envelope = {
        "operations": [
            {
                "op_id": "op_001",
                "op": "replace_para",
                "target_id": "p_000001",
                "items": [{"content_literal": "hi"}],
                "raw": True,
            }
        ]
    }
    with pytest.raises(ValidationError):
        validate_batch(envelope)


def test_schema_rejects_unknown_operation_kind() -> None:
    envelope = {
        "operations": [
            {
                "op_id": "op_001",
                "op": "rewrite_document",
                "target_id": "p_000001",
                "items": [{"content_literal": "hi"}],
            }
        ]
    }
    with pytest.raises(ValidationError):
        validate_batch(envelope)


def test_export_script_produces_current_schema(tmp_path: Path) -> None:
    """Regenerating the schema on disk should not change bytes."""
    on_disk = (SKILL_DIR / "agent_schema.json").read_text(encoding="utf-8")
    fresh = json.dumps(copy.deepcopy(BATCH_SCHEMA), indent=2, ensure_ascii=False) + "\n"
    assert on_disk == fresh
