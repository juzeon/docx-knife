"""Tests for the atomic batch executor (Phase 6)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st
from lxml import etree

from docx_knife import (
    BatchOperationError,
    Document,
    EditOperation,
    EditResult,
    ValidationError,
)
from docx_knife.batch import BatchExecutor

from . import _fixtures


def _texts(doc: Document) -> list[str]:
    return [p.text or "" for p in doc.list_paragraphs(max_chars=0).paragraphs]


def _ids(doc: Document) -> list[str]:
    return [p.id for p in doc.list_paragraphs().paragraphs]


def _root_bytes(doc: Document) -> bytes:
    return etree.tostring(doc._root)


# ---------------------------------------------------------------- happy path


def test_full_success_batch_returns_results_in_input_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        result = doc.batch_edit([
            EditOperation.insert_para_after(
                op_id="op-insert",
                target_id=ids[0],
                items=["A1"],
            ),
            EditOperation.replace_text(
                op_id="op-replace-text",
                paragraph_id=ids[1],
                find="B",
                replacement="B2",
            ),
            EditOperation.replace_para(
                op_id="op-replace-para",
                target_id=ids[2],
                items=["C1", "C2"],
            ),
            EditOperation.delete_para(op_id="op-delete", target_ids=[ids[3]]),
        ])
        assert isinstance(result, EditResult)
        assert [r.op_id for r in result.results] == [
            "op-insert",
            "op-replace-text",
            "op-replace-para",
            "op-delete",
        ]
        assert all(r.status == "success" for r in result.results)
        insert_result = result.results[0]
        assert insert_result.target_id == ids[0]
        assert len(insert_result.new_ids) == 1
        replace_result = result.results[2]
        assert len(replace_result.new_ids) == 2
        delete_result = result.results[3]
        assert delete_result.target_ids == (ids[3],)
        assert _texts(doc) == ["A", "A1", "B2", "C1", "C2"]


def test_result_correlation_matches_input_ids(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        ops = [
            EditOperation.insert_para_after(op_id=f"op-{i}", target_id=ids[0], items=[f"X{i}"])
            for i in range(3)
        ] + [
            EditOperation.replace_text(
                op_id="op-r", paragraph_id=ids[1], find="B", replacement="BB"
            ),
            EditOperation.delete_para(op_id="op-d", target_ids=[ids[3]]),
        ]
        result = doc.batch_edit(ops)
        assert [r.op_id for r in result.results] == [op.op_id for op in ops]


# ---------------------------------------------------------------- conflicts


def test_two_replace_para_on_same_id_rejected(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.replace_para(op_id="a", target_id=ids[0], items=["X"]),
                EditOperation.replace_para(op_id="b", target_id=ids[0], items=["Y"]),
            ])
        assert excinfo.value.op_id == "b"
        assert _texts(doc) == ["A", "B", "C", "D"]


def test_replace_and_delete_on_same_id_rejected(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.replace_para(op_id="a", target_id=ids[0], items=["X"]),
                EditOperation.delete_para(op_id="b", target_ids=[ids[0]]),
            ])


def test_delete_and_after_insert_on_same_id_rejected(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="a", target_id=ids[0], items=["X"]),
                EditOperation.delete_para(op_id="b", target_ids=[ids[0]]),
            ])


def test_delete_and_before_insert_on_same_id_allowed(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        result = doc.batch_edit([
            EditOperation.insert_para_before(op_id="a", target_id=ids[0], items=["Z1", "Z2"]),
            EditOperation.delete_para(op_id="b", target_ids=[ids[0]]),
        ])
        assert [r.status for r in result.results] == ["success", "success"]
        # A is removed, Z1/Z2 land before its slot.
        assert _texts(doc) == ["Z1", "Z2", "B", "C", "D"]


def test_replace_and_after_insert_on_same_id_allowed(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        # Insertion executes first on the ORIGINAL anchor position, then
        # replacement swaps A itself.
        doc.batch_edit([
            EditOperation.insert_para_after(op_id="a", target_id=ids[0], items=["Y"]),
            EditOperation.replace_para(op_id="b", target_id=ids[0], items=["X"]),
        ])
        assert _texts(doc) == ["X", "Y", "B", "C", "D"]


def test_two_after_inserts_on_same_id_merge_items_in_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        doc.batch_edit([
            EditOperation.insert_para_after(op_id="a", target_id=ids[0], items=["X", "Y"]),
            EditOperation.insert_para_after(op_id="b", target_id=ids[0], items=["Z"]),
        ])
        assert _texts(doc) == ["A", "X", "Y", "Z", "B", "C", "D"]


def test_two_before_inserts_on_same_id_merge_items_in_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        doc.batch_edit([
            EditOperation.insert_para_before(op_id="a", target_id=ids[3], items=["X"]),
            EditOperation.insert_para_before(op_id="b", target_id=ids[3], items=["Y", "Z"]),
        ])
        assert _texts(doc) == ["A", "B", "C", "X", "Y", "Z", "D"]


# ---------------------------------------------------------------- prevalidation


def test_prevalidation_blocks_all_ops_when_later_op_invalid(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        before = _root_bytes(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="ok", target_id=ids[0], items=["X"]),
                EditOperation.replace_para(op_id="bad", target_id="p_ZZZZZZ", items=["Y"]),
            ])
        assert excinfo.value.operation_index == 1
        assert excinfo.value.op_id == "bad"
        # None of the operations mutated the DOM.
        assert _root_bytes(doc) == before
        assert doc.change_log() == []


def test_empty_operations_rejected(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        with pytest.raises(ValidationError) as excinfo:
            doc.batch_edit([])
        assert excinfo.value.failed_check == "nonempty"


def test_envelope_schema_validation(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc, pytest.raises(ValidationError):
        doc.batch_edit(
            [],
            envelope={"operations": [{"op_id": "x", "op": "unknown_op"}]},
        )


def test_invalid_regex_selector_rejected(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.replace_text(
                    op_id="bad",
                    paragraph_id=ids[0],
                    find={"pattern": "(unclosed", "regex": True},
                    replacement="X",
                ),
            ])
        assert excinfo.value.op_id == "bad"


# ---------------------------------------------------------------- rollback


def test_mid_batch_failure_rolls_back_dom_manifest_and_log(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        before_xml = _root_bytes(doc)
        before_ids = doc._manifest.ordered_ids()
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="op1", target_id=ids[0], items=["X"]),
                EditOperation.replace_text(
                    op_id="op2",
                    paragraph_id=ids[1],
                    find="NONEXISTENT",
                    replacement="Y",
                ),
            ])
        assert excinfo.value.operation_index == 1
        assert excinfo.value.rolled_back is True
        assert excinfo.value.__cause__ is not None
        assert _root_bytes(doc) == before_xml
        # Manifest ids are exactly the same set as before (in order).
        assert doc._manifest.ordered_ids() == before_ids
        # Change log contains exactly one failed audit event.
        log = doc.change_log()
        assert len(log) == 1
        assert log[0]["status"] == "failed"
        assert log[0]["rolled_back"] is True
        assert log[0]["error_type"] == "TextNotFoundError"


def test_content_ref_failure_rolls_back(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        before_xml = _root_bytes(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(
                    op_id="op1",
                    target_id=ids[0],
                    items=[
                        {
                            "content_ref": {
                                "type": "file",
                                "path": "does_not_exist.txt",
                            }
                        }
                    ],
                ),
            ])
        assert excinfo.value.operation_index == 0
        assert excinfo.value.rolled_back is True
        assert type(excinfo.value.__cause__).__name__ == "InvalidContentError"
        assert _root_bytes(doc) == before_xml


def test_validation_failure_rolls_back(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")

    def _corrupt(
        self: BatchExecutor,
        snapshot: Any,
        pre_ids: Any,
        groups: Any,
        expected_delta: int,
    ) -> None:
        raise ValidationError(
            stage="commit",
            checks=("untouched_fidelity",),
            failed_check="untouched_fidelity",
        )

    monkeypatch.setattr(BatchExecutor, "_precommit_validate", _corrupt)
    with Document.open(src) as doc:
        ids = _ids(doc)
        before_xml = _root_bytes(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="op1", target_id=ids[0], items=["X"]),
            ])
        assert type(excinfo.value.__cause__).__name__ == "ValidationError"
        assert _root_bytes(doc) == before_xml


def test_untouched_fidelity_catches_untracked_mutation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    original = BatchExecutor._execute_plan

    def sneaky(
        self: BatchExecutor,
        groups: dict[str, Any],
        order: list[str],
        results_by_seq: dict[int, Any],
        pending_log: list[dict[str, Any]],
    ) -> int:
        delta = original(self, groups, order, results_by_seq, pending_log)
        # Mutate an untouched paragraph after execution but before validation.
        untouched_id = next(tid for tid in self._doc._manifest.ordered_ids() if tid not in groups)
        node = self._doc._manifest.resolve(untouched_id)
        # Corrupt a <w:t> inside this paragraph.
        for t in node.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
            t.text = "corrupted"
            break
        return delta

    monkeypatch.setattr(BatchExecutor, "_execute_plan", sneaky)
    with Document.open(src) as doc:
        ids = _ids(doc)
        before_xml = _root_bytes(doc)
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="op1", target_id=ids[0], items=["X"]),
            ])
        assert type(excinfo.value.__cause__).__name__ == "ValidationError"
        assert _root_bytes(doc) == before_xml


# ---------------------------------------------------------------- audit log


def test_change_log_success_previews_are_truncated(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        long_text = "X" * 500
        doc.batch_edit([
            EditOperation.insert_para_after(op_id="op1", target_id=ids[0], items=[long_text]),
        ])
        log = doc.change_log()
        assert len(log) == 1
        entry = log[0]
        assert entry["status"] == "success"
        after = entry["after"]
        assert isinstance(after, dict)
        previews = after["previews"]
        assert isinstance(previews, list) and previews
        assert all(len(p) <= 80 for p in previews)


def test_failed_batch_records_single_audit_event(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.insert_para_after(op_id="op1", target_id=ids[0], items=["X"]),
                EditOperation.replace_text(
                    op_id="op2",
                    paragraph_id=ids[1],
                    find="ZZZ",
                    replacement="A",
                ),
            ])
        log = doc.change_log()
        assert len(log) == 1
        assert log[0].get("batch") is True


# ---------------------------------------------------------------- property


@given(
    st.lists(
        st.sampled_from(["insert_after", "insert_before", "delete", "replace"]),
        min_size=1,
        max_size=6,
    )
)
@settings(
    max_examples=25,
    deadline=None,
    suppress_health_check=[HealthCheck.function_scoped_fixture],
)
def test_failed_batch_invariance_property(
    tmp_path_factory: pytest.TempPathFactory, kinds: list[str]
) -> None:
    tmp_path = tmp_path_factory.mktemp("batch")
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = _ids(doc)
        before_xml = _root_bytes(doc)
        ops: list[Any] = []
        # Build a batch whose LAST op is guaranteed to fail so we exercise the
        # rollback path from a random earlier prefix of ops.
        seen_ids: set[str] = set()
        for i, kind in enumerate(kinds[:-1]):
            anchor = ids[i % len(ids)]
            if kind == "delete":
                if anchor in seen_ids:
                    continue
                seen_ids.add(anchor)
                ops.append(EditOperation.delete_para(op_id=f"d{i}", target_ids=[anchor]))
            elif kind == "replace":
                if anchor in seen_ids:
                    continue
                seen_ids.add(anchor)
                ops.append(
                    EditOperation.replace_para(op_id=f"r{i}", target_id=anchor, items=[f"R{i}"])
                )
            elif kind == "insert_before":
                if anchor in seen_ids:
                    continue
                ops.append(
                    EditOperation.insert_para_before(
                        op_id=f"ib{i}", target_id=anchor, items=[f"IB{i}"]
                    )
                )
            else:
                if anchor in seen_ids:
                    continue
                ops.append(
                    EditOperation.insert_para_after(
                        op_id=f"ia{i}", target_id=anchor, items=[f"IA{i}"]
                    )
                )
        # Force failure with an unresolvable text op at the tail.
        ops.append(
            EditOperation.replace_text(
                op_id="boom",
                paragraph_id=ids[0],
                find="__NEVER_MATCH__",
                replacement="!",
            )
        )
        with pytest.raises(BatchOperationError):
            doc.batch_edit(ops)
        assert _root_bytes(doc) == before_xml
