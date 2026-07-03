"""Benchmark: :func:`docx_knife.textmap.build_text_map` on a 200-run paragraph."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docx_knife import Document
from docx_knife.textmap import build_text_map

from ._docs import build_multi_run_paragraph

pytestmark = pytest.mark.benchmark

_BUDGET_NS = 50_000_000  # 50 ms


def test_build_text_map_200_runs(tmp_path: Path) -> None:
    src = build_multi_run_paragraph(tmp_path / "runs.docx", 200)
    with Document.open(src) as doc:
        node = doc._manifest.resolve(doc.list_paragraphs().paragraphs[0].id)

        # Warm-up.
        build_text_map(node)

        start = time.perf_counter_ns()
        text_map = build_text_map(node)
        elapsed = time.perf_counter_ns() - start
        print(f"\n[bench] build_text_map(200 runs) = {elapsed / 1e6:.2f} ms")
        # 200 runs x "segNNN" (6 chars) = 1200 characters.
        assert len(text_map.text) == 1200
        assert elapsed < _BUDGET_NS
