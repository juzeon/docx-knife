"""Benchmark: ``Document.open`` on a 500-paragraph docx."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docx_knife import Document

from ._docs import build_many_paragraphs

pytestmark = pytest.mark.benchmark

_BUDGET_NS = 500_000_000  # 500 ms


def test_open_500_paragraphs(tmp_path: Path) -> None:
    src = build_many_paragraphs(tmp_path / "big.docx", 500)
    # Warm-up (populate FS caches, JIT any lxml import machinery).
    with Document.open(src) as warm:
        assert warm.paragraph_count() == 500

    start = time.perf_counter_ns()
    with Document.open(src) as doc:
        count = doc.paragraph_count()
    elapsed = time.perf_counter_ns() - start
    print(f"\n[bench] open(500 paras)+count = {elapsed / 1e6:.2f} ms")
    assert count == 500
    assert elapsed < _BUDGET_NS, f"open budget blown: {elapsed} ns"
