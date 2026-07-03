"""Contract tests for public models."""

from __future__ import annotations

import dataclasses

import pytest

from docx_knife import (
    ContentItem,
    ContentSourceCommand,
    ContentSourceFile,
    ContentSourceJsonPath,
    DeletePara,
    EditOperation,
    InsertParaAfter,
    InvalidContentError,
    Pagination,
    ParagraphInfo,
    ParagraphLocation,
    ReplacePara,
    ReplaceText,
    Selector,
    TableContext,
    validate_batch,
)
from docx_knife.errors import ValidationError


def _location(*, ordinal: int = 0) -> ParagraphLocation:
    return ParagraphLocation(
        part="word/document.xml",
        original_index=ordinal,
        global_ordinal=ordinal,
        table_context=None,
    )


def test_paragraph_info_requires_exactly_one_content() -> None:
    ParagraphInfo(id="p_1", global_ordinal=0, style_id=None, location=_location(), text="hello")
    ParagraphInfo(id="p_1", global_ordinal=0, style_id=None, location=_location(), xml="<w:p/>")
    with pytest.raises(InvalidContentError):
        ParagraphInfo(id="p_1", global_ordinal=0, style_id=None, location=_location())
    with pytest.raises(InvalidContentError):
        ParagraphInfo(
            id="p_1",
            global_ordinal=0,
            style_id=None,
            location=_location(),
            text="hi",
            xml="<w:p/>",
        )


def test_table_context_is_frozen() -> None:
    ctx = TableContext(
        table_index=0,
        row_index=0,
        physical_cell_index=0,
        logical_column_index=0,
        grid_span=1,
        grid_before=0,
        vertical_merge="none",
        nesting_depth=0,
        paragraph_index_in_cell=0,
    )
    with pytest.raises(dataclasses.FrozenInstanceError):
        ctx.table_index = 99  # type: ignore[misc]


def test_pagination_is_frozen() -> None:
    Pagination(start=1, limit=None, returned=0, total=0)


def test_selector_coerce_variants() -> None:
    assert Selector.coerce("foo") == Selector(pattern="foo", regex=False)
    assert Selector.coerce({"pattern": "bar", "regex": True}) == Selector(
        pattern="bar", regex=True
    )
    original = Selector(pattern="baz", regex=False)
    assert Selector.coerce(original) is original


def test_selector_coerce_invalid() -> None:
    with pytest.raises(InvalidContentError):
        Selector.coerce({"regex": True})
    with pytest.raises(InvalidContentError):
        Selector.coerce({"pattern": 12})  # type: ignore[dict-item]
    with pytest.raises(InvalidContentError):
        Selector.coerce(42)  # type: ignore[arg-type]


def test_content_item_of_literal_and_ref() -> None:
    assert ContentItem.of("hi") == ContentItem(content_literal="hi")
    ref = ContentItem.of(
        {"content_ref": {"type": "file", "path": "clauses.txt"}},
    )
    assert isinstance(ref.content_ref, ContentSourceFile)
    assert ref.content_ref.path == "clauses.txt"
    assert ref.content_ref.encoding == "utf-8"


def test_content_item_rejects_dual_and_missing() -> None:
    with pytest.raises(InvalidContentError):
        ContentItem(content_literal="hi", content_ref=ContentSourceFile(path="a.txt"))
    with pytest.raises(InvalidContentError):
        ContentItem()
    with pytest.raises(InvalidContentError):
        ContentItem.of({"content_literal": "a", "content_ref": {"type": "file", "path": "b"}})


def test_content_item_command_ref() -> None:
    item = ContentItem.of(
        {
            "content_ref": {
                "type": "command",
                "argv": ["python", "render.py"],
                "timeout_seconds": 5,
            }
        }
    )
    assert isinstance(item.content_ref, ContentSourceCommand)
    assert item.content_ref.argv == ("python", "render.py")
    assert item.content_ref.timeout_seconds == 5.0


def test_content_item_jsonpath_ref() -> None:
    item = ContentItem.of(
        {"content_ref": {"type": "jsonpath", "source": "data.json", "path": "$.name"}}
    )
    assert isinstance(item.content_ref, ContentSourceJsonPath)


def test_edit_operation_factories() -> None:
    replace = EditOperation.replace_text(
        op_id="op_1",
        paragraph_id="p_1",
        find="a",
        replacement="b",
        occurrence=0,
    )
    assert isinstance(replace, ReplaceText)
    assert replace.find == Selector(pattern="a", regex=False)
    assert replace.content_literal == "b"

    insert = EditOperation.insert_para_after(
        op_id="op_2",
        target_id="p_1",
        items=["one", "two"],
    )
    assert isinstance(insert, InsertParaAfter)
    assert len(insert.items) == 2

    replaced = EditOperation.replace_para(
        op_id="op_3",
        target_id="p_1",
        items=[ContentItem(content_literal="hi")],
    )
    assert isinstance(replaced, ReplacePara)

    delete = EditOperation.delete_para(op_id="op_4", target_ids=["p_1", "p_2"])
    assert isinstance(delete, DeletePara)


def test_delete_para_rejects_duplicates_and_empty() -> None:
    with pytest.raises(InvalidContentError):
        DeletePara(op_id="op", target_ids=())
    with pytest.raises(InvalidContentError):
        DeletePara(op_id="op", target_ids=("p_1", "p_1"))


def test_replace_text_requires_exactly_one_content() -> None:
    with pytest.raises(InvalidContentError):
        ReplaceText(op_id="op", target_id="p_1", find=Selector(pattern="a"))
    with pytest.raises(InvalidContentError):
        ReplaceText(
            op_id="op",
            target_id="p_1",
            find=Selector(pattern="a"),
            content_literal="x",
            content_ref=ContentSourceFile(path="b.txt"),
        )


def test_validate_batch_accepts_minimal_payload() -> None:
    validate_batch(
        {
            "operations": [
                {
                    "op_id": "op_1",
                    "op": "replace_text",
                    "target_id": "p_1",
                    "find": "hello",
                    "content_literal": "world",
                }
            ]
        }
    )


def test_validate_batch_accepts_all_operation_kinds() -> None:
    payload = {
        "operations": [
            {
                "op_id": "op_ins_before",
                "op": "insert_para_before",
                "target_id": "p_1",
                "items": [{"content_literal": "a"}],
            },
            {
                "op_id": "op_ins_after",
                "op": "insert_para_after",
                "target_id": "p_1",
                "items": [
                    {
                        "content_ref": {
                            "type": "file",
                            "path": "clause.txt",
                        }
                    }
                ],
            },
            {
                "op_id": "op_replace",
                "op": "replace_para",
                "target_id": "p_1",
                "items": [{"content_literal": "b"}],
            },
            {
                "op_id": "op_delete",
                "op": "delete_para",
                "target_ids": ["p_2", "p_3"],
            },
            {
                "op_id": "op_dt",
                "op": "delete_text",
                "target_id": "p_1",
                "find": {"pattern": "x", "regex": False},
            },
            {
                "op_id": "op_it_before",
                "op": "insert_text_before",
                "target_id": "p_1",
                "find": "x",
                "content_literal": "y",
            },
            {
                "op_id": "op_it_after",
                "op": "insert_text_after",
                "target_id": "p_1",
                "find": "x",
                "content_ref": {
                    "type": "command",
                    "argv": ["echo", "hi"],
                    "timeout_seconds": 1.5,
                },
            },
        ]
    }
    validate_batch(payload)


def test_validate_batch_rejects_unknown_op() -> None:
    with pytest.raises(ValidationError):
        validate_batch(
            {
                "operations": [
                    {"op_id": "op", "op": "reformat", "target_id": "p_1"},
                ]
            }
        )


def test_validate_batch_rejects_dual_content() -> None:
    with pytest.raises(ValidationError):
        validate_batch(
            {
                "operations": [
                    {
                        "op_id": "op",
                        "op": "replace_text",
                        "target_id": "p_1",
                        "find": "x",
                        "content_literal": "a",
                        "content_ref": {"type": "file", "path": "b.txt"},
                    }
                ]
            }
        )


def test_validate_batch_rejects_empty_operations() -> None:
    with pytest.raises(ValidationError):
        validate_batch({"operations": []})


def test_validate_batch_rejects_duplicate_target_ids_in_delete() -> None:
    with pytest.raises(ValidationError):
        validate_batch(
            {
                "operations": [
                    {
                        "op_id": "op",
                        "op": "delete_para",
                        "target_ids": ["p_1", "p_1"],
                    }
                ]
            }
        )
