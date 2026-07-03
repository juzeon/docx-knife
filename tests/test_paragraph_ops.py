"""Tests for Phase-4 whole-paragraph operations and the fluent ``Paragraph`` API."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_knife import (
    Document,
    InvalidContentError,
    Paragraph,
    ParagraphNotFoundError,
    ValidationError,
)

from . import _fixtures

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _texts(doc: Document) -> list[str]:
    return [p.text or "" for p in doc.list_paragraphs(max_chars=0).paragraphs]


# ------------------------------------------------------------- insert order


def test_insert_after_preserves_item_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        a_id = ids[0]
        new = doc.insert_para_after(a_id, ["B1", "C1"])
        assert len(new) == 2
        assert all(isinstance(p, Paragraph) for p in new)
        assert _texts(doc) == ["A", "B1", "C1", "B", "C", "D"]


def test_insert_before_preserves_item_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        d_id = ids[3]
        new = doc.insert_para_before(d_id, ["B1", "C1"])
        assert [p.id for p in new] == [new[0].id, new[1].id]
        assert _texts(doc) == ["A", "B", "C", "B1", "C1", "D"]


# ------------------------------------------------------------- item expansion


def test_item_expansion_multi_paragraph_and_line_break(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        new = doc.insert_para_after(a_id, ["line1\n\nline2\n\n\nline3"])
        assert len(new) == 3
        # Single \n embedded in a paragraph becomes <w:br/>.
        new2 = doc.insert_para_after(new[-1].id, ["prefix\nmiddle\nsuffix"])
        elem = doc._manifest.resolve(new2[0].id)
        assert elem.findall(f".//{{{_W}}}br"), "single \\n must project to <w:br/>"


def test_formatting_inherits_pPr_and_first_rPr(tmp_path: Path) -> None:
    src = _fixtures.build_formatted(tmp_path / "fmt.docx")
    with Document.open(src) as doc:
        anchor_id = doc.list_paragraphs().paragraphs[0].id
        [new] = doc.insert_para_after(anchor_id, ["Inserted."])
        elem = doc._manifest.resolve(new.id)
        # <w:pPr> should include <w:pStyle w:val="Heading1"/>.
        pstyle = elem.find(f"{{{_W}}}pPr/{{{_W}}}pStyle")
        assert pstyle is not None
        assert pstyle.get(f"{{{_W}}}val") == "Heading1"
        # First run inherited <w:rPr><w:b/></w:rPr>.
        first_run = elem.find(f"{{{_W}}}r")
        assert first_run is not None
        assert first_run.find(f"{{{_W}}}rPr/{{{_W}}}b") is not None


def test_new_paragraph_preserves_boundary_whitespace(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        [new] = doc.insert_para_after(a_id, [" spaced "])
        elem = doc._manifest.resolve(new.id)
        t = elem.find(f"{{{_W}}}r/{{{_W}}}t")
        assert t is not None
        assert t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"


# ---------------------------------------------------------------- replace


def test_replace_para_invalidates_old_id(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        new = doc.replace_para(a_id, ["X1", "X2"])
        assert len(new) == 2
        with pytest.raises(ParagraphNotFoundError):
            doc.get_paragraph(a_id)
        assert _texts(doc) == ["X1", "X2", "B", "C", "D"]


def test_replace_para_warns_on_protected_structures(tmp_path: Path) -> None:
    src = _fixtures.build_hyperlink_para(tmp_path / "hy.docx")
    with Document.open(src) as doc:
        target_id = doc.list_paragraphs().paragraphs[0].id
        result = doc.replace_para(target_id, ["Plain replacement."])
        assert len(result) == 1
        warnings = doc.last_operation_warnings
        assert "hyperlink" in warnings
        assert "run_styling" in warnings


# ---------------------------------------------------------------- delete


def test_delete_para_multiple_targets(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        before = doc.paragraph_count()
        doc.delete_para([ids[1], ids[2]])
        assert doc.paragraph_count() == before - 2
        assert _texts(doc) == ["A", "D"]


def test_delete_para_rejects_empty(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        with pytest.raises(ValidationError) as excinfo:
            doc.delete_para([])
        assert excinfo.value.failed_check == "nonempty"


def test_delete_para_rejects_duplicates(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        first = doc.list_paragraphs().paragraphs[0].id
        with pytest.raises(ValidationError) as excinfo:
            doc.delete_para([first, first])
        assert excinfo.value.failed_check == "unique"


def test_delete_para_rejects_unknown(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        first = doc.list_paragraphs().paragraphs[0].id
        with pytest.raises(ParagraphNotFoundError):
            doc.delete_para([first, "p_zzzzzz"])


def test_delete_reverse_order_keeps_survivors_intact(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        # Deliberately supply out-of-order IDs; deletion should still work.
        doc.delete_para([ids[2], ids[0]])
        assert _texts(doc) == ["B", "D"]


# ---------------------------------------------------------------- raw mode


def test_raw_replace_para_round_trip(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        raw_xml = doc.get_paragraph(a_id, raw=True)
        fragment = raw_xml + (f'<w:p xmlns:w="{_W}"><w:r><w:t>Second raw.</w:t></w:r></w:p>')
        new = doc.replace_para(a_id, [fragment], raw=True)
        assert len(new) == 2
        assert [new[0].read(), new[1].read()] == ["A", "Second raw."]


def test_raw_fragment_rejects_non_wp_top_level(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        target = doc.list_paragraphs().paragraphs[0].id
        with pytest.raises(InvalidContentError):
            doc.replace_para(target, ["<not-a-p/>"], raw=True)


def test_raw_fragment_rejects_mixed_top_level(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        target = doc.list_paragraphs().paragraphs[0].id
        bad = f'<w:p xmlns:w="{_W}"><w:r><w:t>x</w:t></w:r></w:p><foo/>'
        with pytest.raises(InvalidContentError):
            doc.replace_para(target, [bad], raw=True)


def test_raw_fragment_rejects_malformed_xml(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        target = doc.list_paragraphs().paragraphs[0].id
        with pytest.raises(InvalidContentError):
            doc.replace_para(target, ["<w:p>malformed"], raw=True)


# ---------------------------------------------------------------- tables


def test_insert_after_table_paragraph_stays_in_cell(tmp_path: Path) -> None:
    src = _fixtures.build_table_paragraph(tmp_path / "tbl.docx")
    with Document.open(src) as doc:
        # Find the paragraph inside the table cell.
        infos = doc.list_paragraphs().paragraphs
        cell_info = next(info for info in infos if info.location.table_context is not None)
        [new] = doc.insert_para_after(cell_info.id, ["cell-second"])
        # Both the original and the new paragraph must share the same <w:tc> parent.
        original = doc._manifest.resolve(cell_info.id)
        inserted = doc._manifest.resolve(new.id)
        assert original.getparent() is inserted.getparent()
        assert original.getparent().tag == f"{{{_W}}}tc"


# ---------------------------------------------------------------- fluent chain


def test_fluent_chaining_preserves_order(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        anchor = doc.get_paragraph_object(a_id)
        [b_new] = anchor.insert_para_after(["B_new"])
        [c_new] = b_new.insert_para_after(["C_new"])
        [between] = c_new.insert_para_before(["B_mid"])
        assert _texts(doc) == ["A", "B_new", "B_mid", "C_new", "B", "C", "D"]
        # All returned handles remain valid.
        assert anchor.read() == "A"
        assert between.read() == "B_mid"


def test_paragraph_object_after_replace_is_invalid(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        original = doc.get_paragraph_object(a_id)
        original.replace_para(["Fresh."])
        with pytest.raises(ParagraphNotFoundError):
            original.read()


def test_fluent_text_edit_delegates(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        a_id = doc.list_paragraphs().paragraphs[0].id
        handle = doc.get_paragraph_object(a_id)
        handle.replace_text("A", "ZZ")
        assert handle.read() == "ZZ"
        handle.insert_text_before("ZZ", ">")
        handle.insert_text_after("ZZ", "<")
        assert handle.read() == ">ZZ<"
        handle.delete_text(">")
        assert handle.read() == "ZZ<"


# ---------------------------------------------------------------- empty items


def test_insert_para_rejects_empty_items(tmp_path: Path) -> None:
    src = _fixtures.build_abcd(tmp_path / "abcd.docx")
    with Document.open(src) as doc:
        first = doc.list_paragraphs().paragraphs[0].id
        with pytest.raises(InvalidContentError):
            doc.insert_para_after(first, [])
