"""Tests for Phase-2 read/search APIs on ``Document``."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_knife import (
    AmbiguousTextMatchError,
    Document,
    InvalidPatternError,
    ParagraphNotFoundError,
    TextMatch,
    TextNotFoundError,
)

from . import _fixtures

# ---------------------------------------------------------------- listings


def test_paragraph_count_and_style(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        assert doc.paragraph_count() == 3
        items = doc.list_paragraphs().paragraphs
        assert [p.text for p in items] == [
            "First paragraph.",
            "A heading.",
            "Third paragraph.",
        ]
        assert items[1].style_id == "Heading1"
        assert all(p.xml is None for p in items)


def test_list_paragraphs_pagination_and_truncation(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        window = doc.list_paragraphs(start=2, limit=1, max_chars=3)
        assert window.pagination.total == 3
        assert window.pagination.returned == 1
        assert window.pagination.start == 2
        assert window.pagination.limit == 1
        assert [p.text for p in window.paragraphs] == ["A h"]
        # max_chars=0 returns full content
        full = doc.list_paragraphs(max_chars=0)
        assert full.paragraphs[0].text == "First paragraph."


def test_list_paragraphs_raw_returns_xml(tmp_path: Path) -> None:
    src = _fixtures.build_raw_query(tmp_path / "raw.docx")
    with Document.open(src) as doc:
        listing = doc.list_paragraphs(raw=True, max_chars=0)
        item = listing.paragraphs[0]
        assert item.text is None
        assert item.xml is not None
        assert item.xml.startswith("<w:p")
        assert "违约" in item.xml
        assert "责任" in item.xml


def test_get_paragraph_unknown_id(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc, pytest.raises(ParagraphNotFoundError):
        doc.get_paragraph("p_999999")


def test_get_visible_text_visible_and_raw(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        visible = doc.get_visible_text()
        assert visible == "First paragraph.\nA heading.\nThird paragraph."
        raw = doc.get_visible_text(raw=True)
        assert raw.count("</w:p>") == 3
        assert "<?xml" not in raw


# --------------------------------------------------------------- search


def test_grep_paragraphs_finds_hit_beyond_preview(tmp_path: Path) -> None:
    src = _fixtures.build_raw_query(tmp_path / "raw.docx")
    with Document.open(src) as doc:
        result = doc.grep_paragraphs("三十日", max_chars=3)
        assert result.total_matches == 1
        [hit] = result.matches
        full = doc.get_paragraph(hit.paragraph.id)
        assert hit.paragraph.text == full[:3]  # preview truncated to first 3 chars
        start, end = hit.ranges[0]
        assert full[start:end] == "三十日"


def test_grep_paragraphs_regex(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        result = doc.grep_paragraphs(r"paragraph|heading", regex=True)
        assert result.total_matches == 3


def test_find_text_variants(tmp_path: Path) -> None:
    src = _fixtures.build_raw_query(tmp_path / "raw.docx")
    with Document.open(src) as doc:
        assert doc.find_text("nothing here") is None
        [pid] = [p.id for p in doc.list_paragraphs().paragraphs]
        one = doc.find_text("三十日", paragraph_id=pid)
        assert isinstance(one, TextMatch)
        assert one.total_matches == 1
        # Now a paragraph with two matches (append via low-level for the sake of test)
        node = doc._manifest.resolve(pid)
        node[-1].find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t").text += (  # type: ignore[union-attr,operator]
            "三十日额外"
        )
        doc._invalidate_index()
        with pytest.raises(AmbiguousTextMatchError):
            doc.find_text("三十日", paragraph_id=pid)
        second = doc.find_text("三十日", paragraph_id=pid, occurrence=1)
        assert isinstance(second, TextMatch)
        assert second.total_matches == 2
        all_hits = doc.find_text("三十日", paragraph_id=pid, occurrence=-1)
        assert isinstance(all_hits, list)
        assert len(all_hits) == 2
        with pytest.raises(TextNotFoundError):
            doc.find_text("三十日", paragraph_id=pid, occurrence=5)


def test_count_matches_literal_and_regex(tmp_path: Path) -> None:
    src = _fixtures.build_raw_query(tmp_path / "raw.docx")
    with Document.open(src) as doc:
        assert doc.count_matches("三十日") == 1
        assert doc.count_matches(r"违约|三十日", regex=True) == 2


def test_invalid_regex_raises(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc, pytest.raises(InvalidPatternError):
        doc.count_matches("(unclosed", regex=True)


# ------------------------------------------------------------- structure


def test_sdt_paragraphs_excluded_but_counted_globally(tmp_path: Path) -> None:
    src = _fixtures.build_sdt(tmp_path / "sdt.docx")
    with Document.open(src) as doc:
        listing = doc.list_paragraphs(max_chars=0)
        texts = [p.text for p in listing.paragraphs]
        assert texts == ["A", "C"]
        ordinals = [p.global_ordinal for p in listing.paragraphs]
        # SDT-wrapped B keeps ordinal 2, so C is ordinal 3.
        assert ordinals == [1, 3]


def test_nested_table_metadata(tmp_path: Path) -> None:
    src = _fixtures.build_nested_tables(tmp_path / "nested.docx")
    with Document.open(src) as doc:
        by_text = {p.text: p for p in doc.list_paragraphs(max_chars=0).paragraphs}
        assert set(by_text) >= {
            "preamble",
            "inner-a",
            "inner-b",
            "outer-0-0-tail",
            "outer-0-1",
            "outer-1-span",
            "epilogue",
        }
        # inner table paragraph is nested one level deep.
        inner_ctx = by_text["inner-a"].location.table_context
        assert inner_ctx is not None
        assert inner_ctx.nesting_depth == 1
        assert inner_ctx.table_index >= 1

        # outer-1-span sits in a row with gridBefore=1 and a cell with gridSpan=2.
        span_ctx = by_text["outer-1-span"].location.table_context
        assert span_ctx is not None
        assert span_ctx.grid_before == 1
        assert span_ctx.grid_span == 2
        assert span_ctx.logical_column_index == 1  # gridBefore=1 skips column 0
        assert span_ctx.nesting_depth == 0

        # outer-0-1 comes after cell 0 in row 0 with grid_span=1.
        c01 = by_text["outer-0-1"].location.table_context
        assert c01 is not None
        assert c01.physical_cell_index == 1
        assert c01.logical_column_index == 1


def test_paragraph_info_field_exclusivity(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        visible = doc.list_paragraphs().paragraphs[0]
        raw = doc.list_paragraphs(raw=True, max_chars=0).paragraphs[0]
        assert visible.text is not None and visible.xml is None
        assert raw.xml is not None and raw.text is None
