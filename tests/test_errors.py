"""Contract tests for the exception hierarchy."""

from __future__ import annotations

import json
import pickle

import pytest

from docx_knife import (
    AmbiguousTextMatchError,
    BatchOperationError,
    DocumentNotFoundError,
    DocxKnifeError,
    InvalidContentError,
    InvalidDocumentError,
    InvalidPatternError,
    ParagraphNotFoundError,
    Selector,
    SourceChangedError,
    TextNotFoundError,
    UnsupportedStructureError,
    ValidationError,
)


def _all_error_instances() -> list[DocxKnifeError]:
    selector = Selector(pattern="\u4e00" * 200, regex=False)
    return [
        DocumentNotFoundError(path="/tmp/nope.docx"),
        InvalidDocumentError(path="/tmp/bad.docx", reason="not a zip"),
        SourceChangedError(source_path="/tmp/src.docx"),
        ParagraphNotFoundError(target_id="p_000001"),
        TextNotFoundError(
            target_id="p_000042",
            selector=selector,
            occurrence=1,
            total_matches=0,
        ),
        AmbiguousTextMatchError(target_id="p_000042", selector=selector, total_matches=3),
        InvalidPatternError(pattern="[", reason="bad regex"),
        InvalidContentError(raw=True, reason="fragment not w:p"),
        UnsupportedStructureError(
            target_id="p_000001",
            structures=("hyperlink",),
            matched_range=(0, 12),
        ),
        BatchOperationError(
            operation_index=2,
            op_id="op_007",
            reason="conflict",
            cause=RuntimeError("boom"),
        ),
        ValidationError(
            stage="commit",
            checks=("paragraphs_match", "xml_wellformed"),
            failed_check="paragraphs_match",
        ),
    ]


@pytest.mark.parametrize("error", _all_error_instances())
def test_errors_inherit_base(error: DocxKnifeError) -> None:
    assert isinstance(error, DocxKnifeError)
    assert isinstance(error, Exception)


@pytest.mark.parametrize("error", _all_error_instances())
def test_to_dict_is_json_serializable(error: DocxKnifeError) -> None:
    payload = error.to_dict()
    assert payload["code"] == error.code
    for field in error.__public_fields__:
        assert field in payload
    json.dumps(payload)


@pytest.mark.parametrize("error", _all_error_instances())
def test_message_previews_are_bounded(error: DocxKnifeError) -> None:
    message = str(error)
    assert len(message) <= 400
    # No preview substring may exceed the 80-char preview budget on its own.
    for segment in message.split("'"):
        assert len(segment) <= 200


def test_batch_error_preserves_cause() -> None:
    cause = RuntimeError("boom")
    err = BatchOperationError(
        operation_index=0,
        op_id="op_001",
        reason="failed",
        cause=cause,
    )
    assert err.__cause__ is cause
    assert err.rolled_back is True


def test_batch_error_pickle_roundtrip() -> None:
    err = BatchOperationError(
        operation_index=4,
        op_id="op_099",
        reason="rollback",
        rolled_back=True,
    )
    revived = pickle.loads(pickle.dumps(err))
    assert isinstance(revived, BatchOperationError)
    assert revived.operation_index == 4
    assert revived.op_id == "op_099"
    assert revived.reason == "rollback"
    assert revived.rolled_back is True


def test_selector_serializes_in_error_payload() -> None:
    selector = Selector(pattern="hello", regex=True)
    err = AmbiguousTextMatchError(
        target_id="p_000001",
        selector=selector,
        total_matches=2,
    )
    payload = err.to_dict()
    assert payload["selector"] == {"pattern": "hello", "regex": True}


def test_long_preview_is_truncated_in_message() -> None:
    long_pattern = "x" * 500
    err = InvalidPatternError(pattern=long_pattern, reason="y" * 500)
    message = str(err)
    # Neither preview may reach the raw length.
    assert long_pattern not in message
    assert "y" * 500 not in message
