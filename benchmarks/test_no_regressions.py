"""Opt-in regression gate for the docx-knife benchmark suite.

Runs a single-iteration probe and asserts that no operation's p95 has drifted
past 2x the baseline recorded in ``benchmarks/results/baseline.json``.

Skipped by default. Enable by exporting ``RUN_BENCH_REGRESSION=1`` and having a
baseline on disk::

    RUN_BENCH_REGRESSION=1 .venv/bin/pytest benchmarks/test_no_regressions.py
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import pytest

from benchmarks.benchmark_core import RESULTS_PATH, _build_large_docx, _run_suite

_GATE_MULTIPLIER = 2.0


@pytest.mark.skipif(
    not os.environ.get("RUN_BENCH_REGRESSION"),
    reason="RUN_BENCH_REGRESSION not set — regression gate is opt-in",
)
def test_p95_within_2x_of_baseline() -> None:
    if not RESULTS_PATH.exists():
        pytest.skip("no baseline recorded; run benchmarks/benchmark_core.py --baseline")

    baseline = json.loads(RESULTS_PATH.read_text())

    with tempfile.TemporaryDirectory(prefix="docx-knife-bench-") as tmp_str:
        tmp = Path(tmp_str)
        source = tmp / "large.docx"
        _build_large_docx(source, body_paragraphs=5000, table_rows=100, table_cols=4)
        current = _run_suite(source, tmp, iterations=1)

    for name, stats in current.items():
        base = baseline.get(name)
        if base is None:
            continue  # newly added operation — nothing to compare.
        base_p95 = float(base["p95_ms"])
        now_p95 = float(stats["p95_ms"])
        # Ignore sub-millisecond noise floor.
        if base_p95 < 0.5 and now_p95 < 0.5:
            continue
        assert now_p95 <= base_p95 * _GATE_MULTIPLIER, (
            f"performance regression: {name} p95 {now_p95:.3f}ms exceeds "
            f"2x baseline {base_p95:.3f}ms"
        )
