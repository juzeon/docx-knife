"""Benchmark: 50 mixed operations on a 500-paragraph doc plus snapshot cost.

NOTE: :meth:`docx_knife.document.Document.delete_para` re-runs
``_ooxml.iter_editable_paragraphs`` for each call to order deletions in reverse
document order; the batch snapshot (:func:`docx_knife.batch._paragraph_hashes`)
also serializes every paragraph. Both are whole-DOM scans that dominate the
timing for large batches. Recording here for a future perf pass -- production
code is out of scope for phase 8.2.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docx_knife import Document, EditOperation
from docx_knife.batch import BatchExecutor

from ._docs import build_many_paragraphs

pytestmark = pytest.mark.benchmark

_BATCH_BUDGET_NS = 3_000_000_000  # 3 s (very generous)
_SNAPSHOT_BUDGET_NS = 500_000_000  # 500 ms


def _make_batch(ids: list[str]) -> list:
    ops = []
    # 15 insert_para_after, 15 replace_text, 10 insert_para_before,
    # 5 replace_para, 5 delete_para (single-target each) = 50 ops total,
    # each on a distinct anchor to keep conflicts out of scope.
    cursor = 0
    for i in range(15):
        ops.append(
            EditOperation.insert_para_after(op_id=f"ia{i}", target_id=ids[cursor], items=[f"IA{i}"])
        )
        cursor += 1
    for i in range(15):
        ops.append(
            EditOperation.replace_text(
                op_id=f"rt{i}",
                paragraph_id=ids[cursor],
                find="content",
                replacement="stuff",
            )
        )
        cursor += 1
    for i in range(10):
        ops.append(
            EditOperation.insert_para_before(
                op_id=f"ib{i}", target_id=ids[cursor], items=[f"IB{i}"]
            )
        )
        cursor += 1
    for i in range(5):
        ops.append(
            EditOperation.replace_para(
                op_id=f"rp{i}", target_id=ids[cursor], items=[f"RP{i}a", f"RP{i}b"]
            )
        )
        cursor += 1
    for i in range(5):
        ops.append(EditOperation.delete_para(op_id=f"dp{i}", target_ids=[ids[cursor]]))
        cursor += 1
    return ops


def test_batch_50_mixed_ops(tmp_path: Path) -> None:
    src = build_many_paragraphs(tmp_path / "big.docx", 500)
    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        ops = _make_batch(ids)

        start = time.perf_counter_ns()
        result = doc.batch_edit(ops)
        elapsed = time.perf_counter_ns() - start
        print(f"\n[bench] batch_edit(50 mixed ops on 500 paras) = {elapsed / 1e6:.2f} ms")
        assert len(result.results) == 50
        assert elapsed < _BATCH_BUDGET_NS


def test_snapshot_cost_alone(tmp_path: Path) -> None:
    src = build_many_paragraphs(tmp_path / "big.docx", 500)
    with Document.open(src) as doc:
        executor = BatchExecutor(
            doc,
            [
                EditOperation.replace_text(
                    op_id="noop",
                    paragraph_id=doc.list_paragraphs().paragraphs[0].id,
                    find="Paragraph",
                    replacement="Paragraph",
                ),
            ],
        )
        # Warm-up.
        executor._snapshot()

        start = time.perf_counter_ns()
        snapshot = executor._snapshot()
        elapsed = time.perf_counter_ns() - start
        print(f"[bench] snapshot(500 paras) = {elapsed / 1e6:.2f} ms")
        assert len(snapshot.paragraph_hashes) == 500
        assert elapsed < _SNAPSHOT_BUDGET_NS
