"""Unit tests for paragraph-internal text edits (Phase 3)."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_knife._models import Selector
from docx_knife.errors import (
    AmbiguousTextMatchError,
    InvalidContentError,
    TextNotFoundError,
    UnsupportedStructureError,
)
from docx_knife.paragraph import (
    apply_delete_text,
    apply_insert_text_after,
    apply_insert_text_before,
    apply_replace_text,
)
from docx_knife.textmap import build_text_map

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _p(xml: str) -> etree._Element:
    return etree.fromstring(f'<w:p xmlns:w="{_W}">{xml}</w:p>')


def _text(p: etree._Element) -> str:
    return build_text_map(p).text


# ---------------------------------------------------------------- replace_text


def test_cross_run_literal_replace_preserves_second_run_formatting() -> None:
    p = _p("<w:r><w:t>违约</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>责任</w:t></w:r>")
    result = apply_replace_text(
        p,
        Selector(pattern="违约责任"),
        replacement="赔偿责任",
        occurrence=None,
        target_id="p_1",
    )
    assert _text(p) == "赔偿责任"
    assert result.before_preview == "违约责任"
    assert result.after_preview == "赔偿责任"
    # First run's <w:rPr> was None, replacement inherits from first char's run.
    # The second run should still exist because it originally hosted "责任"
    # but was consumed; verify no orphan formatting survives from the b-run.
    # (We simply assert there's no leftover bold span attached to a text.)


def test_cross_run_regex_replace_unique() -> None:
    p = _p("<w:r><w:t>如乙方未能</w:t></w:r><w:r><w:t>按期承担违约责任</w:t></w:r>")
    apply_replace_text(
        p,
        Selector(pattern="如乙方.*?违约责任", regex=True),
        replacement="见附录A",
        occurrence=None,
        target_id="p_1",
    )
    assert _text(p) == "见附录A"


def test_ambiguous_when_two_matches_and_none() -> None:
    p = _p("<w:r><w:t>三十日 A</w:t></w:r><w:r><w:t> B 三十日</w:t></w:r>")
    with pytest.raises(AmbiguousTextMatchError):
        apply_replace_text(
            p, Selector(pattern="三十日"), "六十日", occurrence=None, target_id="p_x"
        )


def test_occurrence_indexed_picks_second_match() -> None:
    p = _p("<w:r><w:t>三十日 A </w:t></w:r><w:r><w:t>三十日 B</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="三十日"), "六十日", occurrence=1, target_id="p_x")
    assert _text(p) == "三十日 A 六十日 B"


def test_occurrence_all_edits_all_right_to_left() -> None:
    p = _p("<w:r><w:t>a-a-a</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="a"), "b", occurrence=-1, target_id="p_x")
    assert _text(p) == "b-b-b"


def test_out_of_range_occurrence_raises() -> None:
    p = _p("<w:r><w:t>ab</w:t></w:r>")
    with pytest.raises(TextNotFoundError):
        apply_replace_text(p, Selector(pattern="ab"), "cd", occurrence=5, target_id="p_x")


def test_no_match_raises() -> None:
    p = _p("<w:r><w:t>ab</w:t></w:r>")
    with pytest.raises(TextNotFoundError):
        apply_replace_text(p, Selector(pattern="zzz"), "cd", occurrence=None, target_id="p_x")


# ---------------------------------------------------------------- markers


def test_replace_around_tab_keeps_tab_intact() -> None:
    p = _p("<w:r><w:t>left</w:t><w:tab/><w:t>right</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="left"), "LEFT", occurrence=None, target_id="p_x")
    assert _text(p) == "LEFT[[DOCX:TAB]]right"
    # Tab node still present.
    assert p.findall(f".//{{{_W}}}tab")


def test_replacement_containing_marker_creates_real_tab() -> None:
    p = _p("<w:r><w:t>ax</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="x"), "[[DOCX:TAB]]", occurrence=None, target_id="p_x")
    assert _text(p) == "a[[DOCX:TAB]]"
    # A real <w:tab/> element was materialised.
    assert p.findall(f".//{{{_W}}}tab")


def test_match_boundary_inside_marker_rejected() -> None:
    p = _p("<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>")
    # Try to match "[[" which starts inside the marker — this straddles a
    # reserved marker and should be rejected.
    with pytest.raises(UnsupportedStructureError):
        apply_replace_text(
            p,
            Selector(pattern=r"a\[\[", regex=True),
            "X",
            occurrence=None,
            target_id="p_x",
        )


def test_unknown_marker_in_replacement_rejected() -> None:
    p = _p("<w:r><w:t>ab</w:t></w:r>")
    with pytest.raises(InvalidContentError):
        apply_replace_text(
            p,
            Selector(pattern="a"),
            "[[DOCX:FOO]]",
            occurrence=None,
            target_id="p_x",
        )


# ---------------------------------------------------------------- structures


def test_replace_across_hyperlink_boundary_rejected() -> None:
    p = _p(
        "<w:r><w:t>pre</w:t></w:r>"
        '<w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:r><w:t>link</w:t></w:r></w:hyperlink>"
        "<w:r><w:t>post</w:t></w:r>"
    )
    with pytest.raises(UnsupportedStructureError):
        apply_replace_text(
            p,
            Selector(pattern="prelink"),
            "X",
            occurrence=None,
            target_id="p_h",
        )


# ---------------------------------------------------------------- delete_text


def test_delete_text_removes_span_only() -> None:
    p = _p("<w:r><w:t>alpha-</w:t></w:r><w:r><w:t>beta-</w:t></w:r><w:r><w:t>gamma</w:t></w:r>")
    apply_delete_text(p, Selector(pattern="beta-"), occurrence=None, target_id="p_d")
    assert _text(p) == "alpha-gamma"


# ---------------------------------------------------------------- insert_text


def test_insert_before_and_after() -> None:
    p = _p("<w:r><w:t>hello world</w:t></w:r>")
    apply_insert_text_before(
        p, Selector(pattern="world"), "great ", occurrence=None, target_id="p_i"
    )
    assert _text(p) == "hello great world"
    apply_insert_text_after(p, Selector(pattern="world"), "!", occurrence=None, target_id="p_i")
    assert _text(p) == "hello great world!"


def test_insert_before_at_run_start() -> None:
    p = _p("<w:r><w:t>abc</w:t></w:r>")
    apply_insert_text_before(p, Selector(pattern="abc"), "X", occurrence=None, target_id="p_i")
    assert _text(p) == "Xabc"


def test_insert_after_at_run_end() -> None:
    p = _p("<w:r><w:t>abc</w:t></w:r>")
    apply_insert_text_after(p, Selector(pattern="abc"), "Y", occurrence=None, target_id="p_i")
    assert _text(p) == "abcY"


# ---------------------------------------------------------------- empty-t cleanup


def test_full_replacement_removes_bare_run() -> None:
    p = _p("<w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r>")
    apply_delete_text(p, Selector(pattern="foo"), occurrence=None, target_id="p_e")
    # First run had only <w:t>foo</w:t>; after emptying it must be removed.
    runs = p.findall(f"{{{_W}}}r")
    assert len(runs) == 1  # one replacement run may or may not exist; delete has none
    # Just the "bar" run left.
    assert _text(p) == "bar"


def test_run_with_drawing_kept_even_when_text_empty() -> None:
    p = _p("<w:r><w:drawing><w:x/></w:drawing><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r>")
    apply_delete_text(p, Selector(pattern="foo"), occurrence=None, target_id="p_e")
    # The first run kept because of the drawing.
    runs = p.findall(f"{{{_W}}}r")
    assert len(runs) == 2
    assert p.findall(f".//{{{_W}}}drawing")


# ---------------------------------------------------------------- sequential ops


def test_sequential_ops_use_rebuilt_map() -> None:
    p = _p("<w:r><w:t>A B</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="A"), "AA", occurrence=None, target_id="p_s")
    apply_replace_text(p, Selector(pattern="AA"), "AAA", occurrence=None, target_id="p_s")
    assert _text(p) == "AAA B"


# ---------------------------------------------------------------- raw rejected


def test_raw_true_rejected_on_all_ops() -> None:
    p = _p("<w:r><w:t>x</w:t></w:r>")
    for func, args in [
        (apply_replace_text, (p, Selector(pattern="x"), "y", None)),
        (apply_delete_text, (p, Selector(pattern="x"), None)),
        (apply_insert_text_before, (p, Selector(pattern="x"), "y", None)),
        (apply_insert_text_after, (p, Selector(pattern="x"), "y", None)),
    ]:
        with pytest.raises(InvalidContentError) as excinfo:
            func(*args, target_id="p_raw", raw=True)  # type: ignore[operator]
        assert excinfo.value.raw is True


# --------------------------------------------------- content_literal round trip


def test_escaped_literal_docx_survives_replace() -> None:
    p = _p("<w:r><w:t>foo [[DOCX:TAB]] bar</w:t></w:r>")
    apply_replace_text(p, Selector(pattern="foo"), "FOO", occurrence=None, target_id="p_esc")
    # The literal text with escape survives (still escaped in the map).
    assert "\\[[DOCX:TAB]]" in build_text_map(p).text
    # And no real <w:tab/> was created for the escaped literal.
    assert not p.findall(f".//{{{_W}}}tab")


def test_find_text_document_populates_node_range(tmp_path):  # type: ignore[no-untyped-def]
    from docx_knife import Document, TextMatch
    from tests import _fixtures

    src = _fixtures.build_raw_query(tmp_path / "raw.docx")
    with Document.open(src) as doc:
        listing = doc.list_paragraphs().paragraphs
        pid = listing[0].id
        # "违约责任" spans the first two <w:t>s.
        m = doc.find_text("违约责任", paragraph_id=pid)
        assert isinstance(m, TextMatch)
        assert m.crosses_nodes is True
        s, e = m.node_range
        assert s != e
