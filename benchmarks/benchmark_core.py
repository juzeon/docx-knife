"""Performance benchmarks for the docx-knife core APIs (Phase 8.2).

Usage
-----

Run the full baseline suite and write the JSON report:

    .venv/bin/python -m benchmarks.benchmark_core --baseline

Run a single-iteration probe (used by ``benchmarks/test_no_regressions.py``):

    .venv/bin/python -m benchmarks.benchmark_core --iterations 1

Output
------

The baseline invocation writes ``benchmarks/results/baseline.json`` and prints
a human-readable summary. ``benchmarks/results/`` is gitignored; the JSON is
runtime output, not a checked-in fixture.

Guarantees enforced during the run
----------------------------------

* Consecutive ``paragraph_count()`` calls must not re-walk the DOM. The second
  call must be at least 5x faster than the first index build; otherwise the
  benchmark aborts.
* No batch operation performs an O(N**2) DOM scan. The suite patches
  ``iter_editable_paragraphs`` with a counter and asserts the per-op call
  count stays below a small constant.
"""

from __future__ import annotations

import argparse
import json
import random
import statistics
import sys
import tempfile
import time
import zipfile
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any, TypeVar

from docx_knife import Document, EditOperation, _ooxml
from docx_knife.textmap import build_text_map

T = TypeVar("T")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

RESULTS_PATH = Path(__file__).parent / "results" / "baseline.json"


# ---------------------------------------------------------------------------
# Fixture generation
# ---------------------------------------------------------------------------


def _paragraph(index: int) -> str:
    return (
        f"<w:p><w:r><w:t>Paragraph number {index} — "
        f"the quick brown fox jumps over target keyword {index}.</w:t></w:r></w:p>"
    )


def _cell(text: str) -> str:
    return f"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def _table(rows: int, cols: int) -> str:
    body = ["<w:tbl>"]
    for r in range(rows):
        body.append("<w:tr>")
        for c in range(cols):
            body.append(_cell(f"row {r} col {c}"))
        body.append("</w:tr>")
    body.append("</w:tbl>")
    return "".join(body)


def _build_large_docx(
    path: Path, *, body_paragraphs: int, table_rows: int, table_cols: int
) -> Path:
    parts: list[str] = []
    for i in range(body_paragraphs):
        parts.append(_paragraph(i))
    parts.append(_table(table_rows, table_cols))
    body = "".join(parts)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}">'
        f"<w:body>{body}</w:body>"
        "</w:document>"
    )
    _rels_ct = "application/vnd.openxmlformats-package.relationships+xml"
    _main_ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f'<Default Extension="rels" ContentType="{_rels_ct}"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/word/document.xml" ContentType="{_main_ct}"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Target="word/document.xml"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/document.xml", document_xml)
    return path


# ---------------------------------------------------------------------------
# Instrumentation
# ---------------------------------------------------------------------------


class _IterCounter:
    """Callable wrapper around ``iter_editable_paragraphs`` that counts calls."""

    def __init__(self, wrapped: Callable[..., Iterator[Any]]) -> None:
        self._wrapped = wrapped
        self.calls = 0

    def __call__(self, *args: Any, **kwargs: Any) -> Iterator[Any]:
        self.calls += 1
        return self._wrapped(*args, **kwargs)


@contextmanager
def _count_iter_calls() -> Iterator[_IterCounter]:
    original = _ooxml.iter_editable_paragraphs
    counter = _IterCounter(original)
    _ooxml.iter_editable_paragraphs = counter  # type: ignore[assignment]
    # The batch executor and document module resolve the symbol at call time
    # through the ``_ooxml`` module, so patching the module attribute is enough.
    try:
        yield counter
    finally:
        _ooxml.iter_editable_paragraphs = original  # type: ignore[assignment]


# ---------------------------------------------------------------------------
# Timing utilities
# ---------------------------------------------------------------------------


def _measure(fn: Callable[[], T], iterations: int) -> tuple[list[float], T]:
    samples: list[float] = []
    last: Any = None
    for _ in range(iterations):
        t0 = time.perf_counter()
        last = fn()
        samples.append((time.perf_counter() - t0) * 1000.0)
    return samples, last


def _summarize(samples: list[float]) -> dict[str, float | int]:
    ordered = sorted(samples)
    return {
        "count": len(samples),
        "mean_ms": statistics.fmean(ordered),
        "p50_ms": ordered[len(ordered) // 2],
        "p95_ms": ordered[min(len(ordered) - 1, max(0, int(round(len(ordered) * 0.95)) - 1))],
        "min_ms": ordered[0],
        "max_ms": ordered[-1],
    }


# ---------------------------------------------------------------------------
# Benchmark scenarios
# ---------------------------------------------------------------------------


def _random_paragraph_ids(doc: Document, k: int, seed: int) -> list[str]:
    rng = random.Random(seed)
    ids = [p.id for p in doc.list_paragraphs().paragraphs]
    return rng.sample(ids, k=min(k, len(ids)))


def _mixed_batch_ops(doc: Document, count: int, seed: int) -> list[Any]:
    rng = random.Random(seed)
    ids = [p.id for p in doc.list_paragraphs().paragraphs]
    if len(ids) < count * 3:
        raise RuntimeError("fixture is too small for mixed batch")

    # Reserve disjoint slices so ops do not conflict on the same anchor.
    insert_anchors = ids[:count]
    replace_anchors = ids[count : 2 * count]
    delete_anchors = ids[2 * count : 3 * count]
    rng.shuffle(insert_anchors)

    ops: list[Any] = []
    for i, anchor in enumerate(insert_anchors[: count // 3]):
        ops.append(
            EditOperation.insert_para_after(
                op_id=f"ins-{i}",
                target_id=anchor,
                items=[f"inserted-{i}"],
            )
        )
    for i, anchor in enumerate(replace_anchors[: count // 3]):
        ops.append(
            EditOperation.replace_text(
                op_id=f"rep-{i}",
                paragraph_id=anchor,
                find="target",
                replacement="TARGET",
            )
        )
    delete_ids = delete_anchors[: count - len(ops)]
    if delete_ids:
        ops.append(EditOperation.delete_para(op_id="del-batch", target_ids=delete_ids))
    return ops


def _run_suite(source: Path, workspace: Path, iterations: int) -> dict[str, dict[str, float | int]]:
    report: dict[str, dict[str, float | int]] = {}

    # --- Document.open (parse + anchor manifest build).
    open_samples: list[float] = []
    for _ in range(iterations):
        t0 = time.perf_counter()
        doc = Document.open(source)
        open_samples.append((time.perf_counter() - t0) * 1000.0)
        doc.close()
    report["Document.open"] = _summarize(open_samples)

    with Document.open(source) as doc:
        # --- paragraph_count(): first call builds index, second is cached.
        t0 = time.perf_counter()
        total = doc.paragraph_count()
        first_ms = (time.perf_counter() - t0) * 1000.0
        # Force cache: warm-up the second-call sample multiple times.
        second_samples: list[float] = []
        for _ in range(iterations * 5):
            t0 = time.perf_counter()
            doc.paragraph_count()
            second_samples.append((time.perf_counter() - t0) * 1000.0)
        report["paragraph_count.first"] = _summarize([first_ms])
        report["paragraph_count.cached"] = _summarize(second_samples)
        # Sanity: cached must be dramatically faster than the initial build.
        # Skip the check when the workload is tiny (both under 0.1 ms) — timer
        # noise dominates.
        if first_ms > 0.5:
            cached_median = statistics.median(second_samples)
            if cached_median > first_ms / 5.0:
                raise RuntimeError(
                    "paragraph_count regression: cached call is not O(1) "
                    f"(first={first_ms:.3f}ms cached_median={cached_median:.3f}ms)"
                )

        # --- list_paragraphs pagination: first page vs last page.
        first_page_samples, _ = _measure(
            lambda: doc.list_paragraphs(start=1, limit=100), iterations
        )
        report["list_paragraphs.first_page"] = _summarize(first_page_samples)

        last_start = max(1, total - 100 + 1)
        last_page_samples, _ = _measure(
            lambda: doc.list_paragraphs(start=last_start, limit=100), iterations
        )
        report["list_paragraphs.last_page"] = _summarize(last_page_samples)

        first_p95 = float(report["list_paragraphs.first_page"]["p95_ms"])
        last_p95 = float(report["list_paragraphs.last_page"]["p95_ms"])
        if first_p95 > 0.5 and last_p95 > first_p95 * 3.0:
            raise RuntimeError(
                "list_paragraphs last-page regression: last p95="
                f"{last_p95:.3f}ms exceeds 3x first p95={first_p95:.3f}ms"
            )

        # --- find_text over the whole document.
        find_samples, _ = _measure(lambda: doc.find_text("target", occurrence=-1), iterations)
        report["find_text.occurrence_all"] = _summarize(find_samples)

        # --- build_text_map on 100 random paragraphs (single measurement per
        # iteration averages 100 builds).
        sample_ids = _random_paragraph_ids(doc, 100, seed=1234)
        nodes = [doc._manifest.resolve(pid) for pid in sample_ids]

        def build_all() -> int:
            for node in nodes:
                build_text_map(node)
            return len(nodes)

        tm_samples, _ = _measure(build_all, iterations)
        # Report per-paragraph average by dividing.
        per_para = [s / len(nodes) for s in tm_samples]
        report["build_text_map.per_paragraph"] = _summarize(per_para)

    # --- Batch edit (100 mixed ops) — measured on a fresh doc each iteration
    #     because ops mutate state.
    batch_samples: list[float] = []
    per_op_iter_counts: list[float] = []
    for iteration in range(iterations):
        with Document.open(source) as doc:
            ops = _mixed_batch_ops(doc, count=100, seed=42 + iteration)
            with _count_iter_calls() as counter:
                t0 = time.perf_counter()
                doc.batch_edit(ops)
                batch_samples.append((time.perf_counter() - t0) * 1000.0)
            per_op_iter_counts.append(counter.calls / len(ops))
    report["batch_edit.100_mixed_ops"] = _summarize(batch_samples)

    max_per_op = max(per_op_iter_counts)
    if max_per_op > 3.0:
        raise RuntimeError(
            "batch_edit performs an O(N**2) scan: "
            f"iter_editable_paragraphs called {max_per_op:.2f}x per op "
            "(expected ≤ 3)"
        )

    # --- Full save round-trip.
    save_samples: list[float] = []
    for iteration in range(iterations):
        dest = workspace / f"save-{iteration}.docx"
        with Document.open(source) as doc:
            t0 = time.perf_counter()
            doc.save(dest)
            save_samples.append((time.perf_counter() - t0) * 1000.0)
    report["Document.save"] = _summarize(save_samples)

    return report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _print_summary(report: dict[str, dict[str, float | int]]) -> None:
    width = max(len(name) for name in report)
    print(f"{'operation':<{width}}  {'p50 ms':>10}  {'p95 ms':>10}  {'mean ms':>10}  count")
    print("-" * (width + 45))
    for name in sorted(report):
        stats = report[name]
        print(
            f"{name:<{width}}  "
            f"{float(stats['p50_ms']):>10.3f}  "
            f"{float(stats['p95_ms']):>10.3f}  "
            f"{float(stats['mean_ms']):>10.3f}  "
            f"{int(stats['count']):>5}"
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="docx-knife performance suite")
    parser.add_argument(
        "--baseline",
        action="store_true",
        help="Write results to benchmarks/results/baseline.json (default: 5 iterations).",
    )
    parser.add_argument("--iterations", type=int, default=None)
    parser.add_argument(
        "--paragraphs",
        type=int,
        default=5000,
        help="Number of body paragraphs in the synthetic fixture (default 5000).",
    )
    parser.add_argument(
        "--table-rows",
        type=int,
        default=100,
        help="Number of table rows appended after the body (default 100).",
    )
    parser.add_argument(
        "--table-cols",
        type=int,
        default=4,
        help="Number of columns per table row (default 4).",
    )
    args = parser.parse_args(argv)

    iterations = args.iterations if args.iterations is not None else 5
    if iterations < 1:
        raise SystemExit("iterations must be >= 1")

    with tempfile.TemporaryDirectory(prefix="docx-knife-bench-") as tmp_str:
        tmp = Path(tmp_str)
        source = tmp / "large.docx"
        _build_large_docx(
            source,
            body_paragraphs=args.paragraphs,
            table_rows=args.table_rows,
            table_cols=args.table_cols,
        )
        report = _run_suite(source, tmp, iterations)

    _print_summary(report)

    if args.baseline:
        RESULTS_PATH.parent.mkdir(parents=True, exist_ok=True)
        RESULTS_PATH.write_text(json.dumps(report, indent=2, sort_keys=True))
        print(f"\nBaseline written to {RESULTS_PATH}")
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entrypoint
    sys.exit(main())
