"""Unit tests for :mod:`docx_knife.textmap`."""

from __future__ import annotations

from lxml import etree

from docx_knife._models import Selector
from docx_knife.errors import (
    InvalidContentError,
    InvalidPatternError,
    UnsupportedStructureError,
)
from docx_knife.textmap import (
    build_text_map,
    compile_selector,
    evaluate_capability,
    range_hits_atomic,
    restore_markers,
    select_matches,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _p(xml: str) -> etree._Element:
    return etree.fromstring(f'<w:p xmlns:w="{_W}">{xml}</w:p>')


def test_adjacent_runs_form_continuous_text() -> None:
    p = _p("<w:r><w:t>违约</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>责任</w:t></w:r>")
    m = build_text_map(p)
    assert m.text == "违约责任"
    # Every char maps to a run; first two share one run, last two share another.
    assert m.positions[0].run_ref is not m.positions[2].run_ref
    assert m.positions[0].run_ref is m.positions[1].run_ref
    assert m.positions[2].run_ref is m.positions[3].run_ref


def test_tab_break_cr_projection() -> None:
    p = _p(
        "<w:r><w:t>a</w:t>"
        '<w:tab/><w:br/><w:br w:type="page"/>'
        '<w:br w:type="column"/><w:cr/>'
        "<w:t>b</w:t></w:r>"
    )
    m = build_text_map(p)
    assert m.text == (
        "a[[DOCX:TAB]][[DOCX:LINE_BREAK]][[DOCX:PAGE_BREAK]][[DOCX:COLUMN_BREAK]][[DOCX:CR]]b"
    )
    # Every marker span should be recorded as atomic.
    tokens = [
        "[[DOCX:TAB]]",
        "[[DOCX:LINE_BREAK]]",
        "[[DOCX:PAGE_BREAK]]",
        "[[DOCX:COLUMN_BREAK]]",
        "[[DOCX:CR]]",
    ]
    covered = 0
    for tok in tokens:
        idx = m.text.index(tok, covered)
        assert (idx, idx + len(tok)) in m.atomic_ranges
        covered = idx + len(tok)


def test_ins_included_del_excluded() -> None:
    p = _p(
        "<w:r><w:t>alpha</w:t></w:r>"
        "<w:ins><w:r><w:t>-INS-</w:t></w:r></w:ins>"
        "<w:del><w:r><w:delText>GONE</w:delText></w:r></w:del>"
        "<w:r><w:t>omega</w:t></w:r>"
    )
    m = build_text_map(p)
    assert "INS" in m.text
    assert "GONE" not in m.text
    # Structure metadata for the -INS- portion should tag w:ins
    ins_start = m.text.index("-INS-")
    tags = m.structures_in_range(ins_start, ins_start + 5)
    assert "w:ins" in tags


def test_literal_docx_escape_is_reversible() -> None:
    p = _p("<w:r><w:t>before [[DOCX:TAB]] after</w:t></w:r>")
    m = build_text_map(p)
    assert "\\[[DOCX:TAB]]" in m.text
    # And the escape span is atomic to protect it from boundary matches.
    idx = m.text.index("\\[[DOCX:TAB]]")
    # Some sub-range boundary should be rejected.
    assert range_hits_atomic(m.atomic_ranges, idx + 2, idx + 5)


def test_structures_in_range_reports_hyperlink() -> None:
    p = _p(
        '<w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:r><w:t>docs</w:t></w:r></w:hyperlink>"
        "<w:r><w:t>tail</w:t></w:r>"
    )
    m = build_text_map(p)
    assert m.text == "docstail"
    assert m.structures_in_range(0, 4) == ("w:hyperlink",)
    assert m.structures_in_range(4, 8) == ()


def test_bookmark_point_markers_do_not_pollute() -> None:
    p = _p(
        '<w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>xy</w:t></w:r><w:bookmarkEnd w:id="0"/>'
    )
    m = build_text_map(p)
    assert m.text == "xy"
    # bookmarkStart/End sit as siblings — they never colour characters.
    assert m.structures_in_range(0, 2) == ()


def test_restore_markers_round_trip() -> None:
    segments = restore_markers("a[[DOCX:TAB]]b\\[[DOCX:literal")
    assert segments == [
        ("text", "a"),
        ("marker", "TAB"),
        ("text", "b[[DOCX:literal"),
    ]


def test_restore_markers_rejects_unknown() -> None:
    try:
        restore_markers("[[DOCX:FOO]]")
    except InvalidContentError as exc:
        assert "unknown" in exc.reason
    else:
        raise AssertionError("expected InvalidContentError")


def test_restore_markers_rejects_unterminated() -> None:
    try:
        restore_markers("[[DOCX:TAB")
    except InvalidContentError as exc:
        assert "unterminated" in exc.reason
    else:
        raise AssertionError("expected InvalidContentError")


def test_evaluate_capability_rejects_hyperlink() -> None:
    try:
        evaluate_capability(("w:hyperlink",), (0, 4), target_id="p_x")
    except UnsupportedStructureError as exc:
        assert exc.structures == ("w:hyperlink",)
        assert exc.matched_range == (0, 4)
    else:
        raise AssertionError("expected UnsupportedStructureError")


def test_evaluate_capability_allows_bookmark_and_ins() -> None:
    # bookmarkStart is 'preserve', w:ins is 'allow' — no exception.
    evaluate_capability(("w:bookmarkStart", "w:ins"), (0, 2), target_id="p")


def test_compile_selector_literal() -> None:
    matcher = compile_selector(Selector(pattern="ab"))
    assert matcher("ababab") == [(0, 2), (2, 4), (4, 6)]


def test_compile_selector_regex_dotall_and_zero_length() -> None:
    matcher = compile_selector(Selector(pattern="a.*?b", regex=True))
    assert matcher("a\nb-a\nb") == [(0, 3), (4, 7)]

    empty = compile_selector(Selector(pattern="a*", regex=True))
    # zero-length matches skipped
    assert empty("bbb") == []


def test_compile_selector_invalid() -> None:
    try:
        compile_selector(Selector(pattern="(", regex=True))
    except InvalidPatternError:
        pass
    else:
        raise AssertionError("expected InvalidPatternError")

    try:
        compile_selector(Selector(pattern="", regex=False))
    except InvalidPatternError:
        pass
    else:
        raise AssertionError("expected InvalidPatternError for empty literal")


def test_select_matches_shapes() -> None:
    assert select_matches([], None) == []
    assert select_matches([(0, 1)], None) == [(0, 1)]
    # ambiguous returns full list; caller must check
    assert select_matches([(0, 1), (2, 3)], None) == [(0, 1), (2, 3)]
    # -1 reverses
    assert select_matches([(0, 1), (2, 3)], -1) == [(2, 3), (0, 1)]
    # index in range
    assert select_matches([(0, 1), (2, 3)], 1) == [(2, 3)]
    # out of range
    assert select_matches([(0, 1)], 5) == []


def test_node_ordinals_track_document_order() -> None:
    p = _p("<w:r><w:t>ab</w:t><w:tab/><w:t>cd</w:t></w:r>")
    m = build_text_map(p)
    ords = m.node_ordinals
    # Three contributing nodes (two <w:t>, one <w:tab>) — ordinals 0,1,2.
    assert sorted(ords.values()) == [0, 1, 2]


# --------------------------------------------------------------- property tests


def test_random_run_split_preserves_text() -> None:
    from hypothesis import given, settings
    from hypothesis import strategies as st

    alphabet = st.characters(min_codepoint=0x20, max_codepoint=0x7A)

    @given(
        pieces=st.lists(
            st.text(alphabet=alphabet, min_size=1, max_size=5),
            min_size=1,
            max_size=6,
        )
    )
    @settings(max_examples=50, deadline=None)
    def inner(pieces: list[str]) -> None:
        # Avoid the escape-triggering '[' at start of a piece for simpler assertion.
        pieces = [p.replace("[", "x") for p in pieces]
        runs = "".join(f"<w:r><w:t>{_xml_escape(piece)}</w:t></w:r>" for piece in pieces)
        p = _p(runs)
        m = build_text_map(p)
        assert m.text == "".join(pieces)
        assert len(m.positions) == len(m.text)

    inner()


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
