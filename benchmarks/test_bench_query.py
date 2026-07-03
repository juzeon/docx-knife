"""Benchmark: ``list_paragraphs``, ``grep_paragraphs`` (regex), and paginated
``find_text`` on a 1000-paragraph document.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docx_knife import Document

from ._docs import build_many_paragraphs

pytestmark = pytest.mark.benchmark

_LIST_BUDGET_NS = 200_000_000  # 200 ms
_GREP_BUDGET_NS = 400_000_000  # 400 ms
_FIND_BUDGET_NS = 400_000_000  # 400 ms


def test_query_apis_on_1000_paragraphs(tmp_path: Path) -> None:
    src = build_many_paragraphs(tmp_path / "big.docx", 1000)
    with Document.open(src) as doc:
        assert doc.paragraph_count() == 1000

        start = time.perf_counter_ns()
        page1 = doc.list_paragraphs(start=1, limit=200, max_chars=80)
        page2 = doc.list_paragraphs(start=201, limit=200, max_chars=80)
        elapsed_list = time.perf_counter_ns() - start
        print(f"\n[bench] list_paragraphs 2x200 windows = {elapsed_list / 1e6:.2f} ms")
        assert len(page1.paragraphs) == 200 and len(page2.paragraphs) == 200
        assert elapsed_list < _LIST_BUDGET_NS

        start = time.perf_counter_ns()
        grep = doc.grep_paragraphs(r"Paragraph \d+ content", regex=True)
        elapsed_grep = time.perf_counter_ns() - start
        print(f"[bench] grep_paragraphs regex = {elapsed_grep / 1e6:.2f} ms")
        assert grep.total_matches >= 1000
        assert elapsed_grep < _GREP_BUDGET_NS

        # Iterate every occurrence of the keyword via find_text pagination.
        start = time.perf_counter_ns()
        results = doc.find_text("KEYWORD", occurrence=-1)
        elapsed_find = time.perf_counter_ns() - start
        print(f"[bench] find_text(occurrence=-1) = {elapsed_find / 1e6:.2f} ms")
        assert isinstance(results, list)
        assert len(results) >= 100
        assert elapsed_find < _FIND_BUDGET_NS
