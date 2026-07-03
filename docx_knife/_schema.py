"""JSON schema for the LLM-facing (``raw=false``) batch envelope."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Final

from jsonschema import Draft202012Validator
from jsonschema import ValidationError as _JsonSchemaValidationError

from .errors import ValidationError

_SELECTOR: Final[dict[str, Any]] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["pattern"],
    "properties": {
        "pattern": {"type": "string", "minLength": 1},
        "regex": {"type": "boolean"},
    },
}

_FIND: Final[dict[str, Any]] = {
    "oneOf": [
        {"type": "string", "minLength": 1},
        _SELECTOR,
    ]
}

_CONTENT_REF: Final[dict[str, Any]] = {
    "oneOf": [
        {
            "type": "object",
            "additionalProperties": False,
            "required": ["type", "source", "path"],
            "properties": {
                "type": {"const": "jsonpath"},
                "source": {"type": "string", "minLength": 1},
                "path": {"type": "string", "minLength": 1},
            },
        },
        {
            "type": "object",
            "additionalProperties": False,
            "required": ["type", "path"],
            "properties": {
                "type": {"const": "file"},
                "path": {"type": "string", "minLength": 1},
                "encoding": {"type": "string", "minLength": 1},
            },
        },
        {
            "type": "object",
            "additionalProperties": False,
            "required": ["type", "argv"],
            "properties": {
                "type": {"const": "command"},
                "argv": {
                    "type": "array",
                    "items": {"type": "string"},
                    "minItems": 1,
                },
                "timeout_seconds": {"type": "number", "exclusiveMinimum": 0},
                "env": {
                    "type": "object",
                    "additionalProperties": {"type": "string"},
                },
                "cwd": {"type": "string"},
            },
        },
    ]
}

_CONTENT_ITEM: Final[dict[str, Any]] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "content_literal": {"type": "string"},
        "content_ref": _CONTENT_REF,
    },
    "oneOf": [
        {"required": ["content_literal"], "not": {"required": ["content_ref"]}},
        {"required": ["content_ref"], "not": {"required": ["content_literal"]}},
    ],
}

_ITEMS: Final[dict[str, Any]] = {
    "type": "array",
    "minItems": 1,
    "items": _CONTENT_ITEM,
}

_OP_ID: Final[dict[str, Any]] = {"type": "string", "minLength": 1}
_TARGET_ID: Final[dict[str, Any]] = {"type": "string", "minLength": 1}


def _paragraph_op(op: str) -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "required": ["op_id", "op", "target_id", "items"],
        "properties": {
            "op_id": _OP_ID,
            "op": {"const": op},
            "target_id": _TARGET_ID,
            "items": _ITEMS,
        },
    }


_DELETE_PARA: Final[dict[str, Any]] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["op_id", "op", "target_ids"],
    "properties": {
        "op_id": _OP_ID,
        "op": {"const": "delete_para"},
        "target_ids": {
            "type": "array",
            "minItems": 1,
            "uniqueItems": True,
            "items": _TARGET_ID,
        },
    },
}


def _text_op(op: str, *, needs_content: bool) -> dict[str, Any]:
    properties: dict[str, Any] = {
        "op_id": _OP_ID,
        "op": {"const": op},
        "target_id": _TARGET_ID,
        "find": _FIND,
        "occurrence": {"type": "integer", "minimum": -1},
    }
    required = ["op_id", "op", "target_id", "find"]
    if needs_content:
        properties["content_literal"] = {"type": "string"}
        properties["content_ref"] = _CONTENT_REF
        return {
            "type": "object",
            "additionalProperties": False,
            "required": required,
            "properties": properties,
            "oneOf": [
                {"required": ["content_literal"], "not": {"required": ["content_ref"]}},
                {"required": ["content_ref"], "not": {"required": ["content_literal"]}},
            ],
        }
    return {
        "type": "object",
        "additionalProperties": False,
        "required": required,
        "properties": properties,
    }


BATCH_SCHEMA: Final[dict[str, Any]] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "docx-knife batch",
    "type": "object",
    "additionalProperties": False,
    "required": ["operations"],
    "properties": {
        "operations": {
            "type": "array",
            "minItems": 1,
            "items": {
                "oneOf": [
                    _paragraph_op("insert_para_before"),
                    _paragraph_op("insert_para_after"),
                    _paragraph_op("replace_para"),
                    _DELETE_PARA,
                    _text_op("replace_text", needs_content=True),
                    _text_op("delete_text", needs_content=False),
                    _text_op("insert_text_before", needs_content=True),
                    _text_op("insert_text_after", needs_content=True),
                ]
            },
        }
    },
}


_VALIDATOR: Final[Draft202012Validator] = Draft202012Validator(BATCH_SCHEMA)


def validate_batch(payload: Mapping[str, Any]) -> None:
    """Validate a batch envelope against :data:`BATCH_SCHEMA`.

    Raises :class:`docx_knife.errors.ValidationError` on failure.
    """
    errors = sorted(_VALIDATOR.iter_errors(payload), key=lambda err: err.path)
    if not errors:
        return
    first = errors[0]
    raise ValidationError(
        stage="batch_schema",
        checks=tuple(_format_error(err) for err in errors),
        failed_check=_format_error(first),
    )


def _format_error(err: _JsonSchemaValidationError) -> str:
    location = "/".join(str(part) for part in err.absolute_path) or "<root>"
    return f"{location}: {err.message}"


__all__ = ["BATCH_SCHEMA", "validate_batch"]
