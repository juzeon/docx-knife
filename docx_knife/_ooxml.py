"""OOXML namespace constants, secure lxml parser, and traversal helpers.

Phase 2 exposes only the pieces the query/anchor layer needs. Later phases
(``TextMap``, paragraph editing) will build on top of these primitives.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Final

from lxml import etree

from ._models import ParagraphLocation, TableContext, VerticalMerge

W_NS: Final[str] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS: Final[str] = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_NS: Final[str] = "http://www.w3.org/XML/1998/namespace"

NSMAP: Final[dict[str, str]] = {
    "w": W_NS,
    "w14": W14_NS,
}


def qn(tag: str) -> str:
    """Expand a namespace-prefixed tag (``w:p`` -> ``{...}p``)."""
    prefix, _, local = tag.partition(":")
    if not local:
        return tag
    ns = NSMAP.get(prefix)
    if ns is None:
        raise KeyError(f"unknown xml namespace prefix: {prefix!r}")
    return f"{{{ns}}}{local}"


P_TAG: Final[str] = qn("w:p")
TBL_TAG: Final[str] = qn("w:tbl")
TR_TAG: Final[str] = qn("w:tr")
TC_TAG: Final[str] = qn("w:tc")
SDT_TAG: Final[str] = qn("w:sdt")
SDT_CONTENT_TAG: Final[str] = qn("w:sdtContent")
T_TAG: Final[str] = qn("w:t")
R_TAG: Final[str] = qn("w:r")
TAB_TAG: Final[str] = qn("w:tab")
BR_TAG: Final[str] = qn("w:br")
CR_TAG: Final[str] = qn("w:cr")
INS_TAG: Final[str] = qn("w:ins")
DEL_TAG: Final[str] = qn("w:del")
DELTEXT_TAG: Final[str] = qn("w:delText")
PPR_TAG: Final[str] = qn("w:pPr")
PSTYLE_TAG: Final[str] = qn("w:pStyle")
TCPR_TAG: Final[str] = qn("w:tcPr")
TRPR_TAG: Final[str] = qn("w:trPr")
GRID_SPAN_TAG: Final[str] = qn("w:gridSpan")
GRID_BEFORE_TAG: Final[str] = qn("w:gridBefore")
VMERGE_TAG: Final[str] = qn("w:vMerge")
W_VAL_ATTR: Final[str] = qn("w:val")
W14_PARA_ID_ATTR: Final[str] = qn("w14:paraId")


def build_secure_parser() -> etree.XMLParser:
    """Return an lxml parser that refuses entities, DTDs, and network access."""
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        huge_tree=False,
        load_dtd=False,
        dtd_validation=False,
    )


def is_sdt_descendant(elem: etree._Element) -> bool:
    """True iff any ancestor of ``elem`` is ``w:sdt`` or ``w:sdtContent``."""
    parent = elem.getparent()
    while parent is not None:
        if parent.tag in (SDT_TAG, SDT_CONTENT_TAG):
            return True
        parent = parent.getparent()
    return False


def _innermost_ancestor(elem: etree._Element, tag: str) -> etree._Element | None:
    parent = elem.getparent()
    while parent is not None:
        if parent.tag == tag:
            return parent
        parent = parent.getparent()
    return None


def _int_val(elem: etree._Element | None, default: int) -> int:
    if elem is None:
        return default
    raw = elem.get(W_VAL_ATTR)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _grid_span(tc: etree._Element) -> int:
    tcpr = tc.find(TCPR_TAG)
    if tcpr is None:
        return 1
    gs = tcpr.find(GRID_SPAN_TAG)
    return max(1, _int_val(gs, 1))


def _grid_before(tr: etree._Element) -> int:
    trpr = tr.find(TRPR_TAG)
    if trpr is None:
        return 0
    gb = trpr.find(GRID_BEFORE_TAG)
    return max(0, _int_val(gb, 0))


def _vertical_merge(tc: etree._Element) -> VerticalMerge:
    tcpr = tc.find(TCPR_TAG)
    if tcpr is None:
        return "none"
    vmerge = tcpr.find(VMERGE_TAG)
    if vmerge is None:
        return "none"
    val = vmerge.get(W_VAL_ATTR)
    if val == "restart":
        return "restart"
    return "continue"


def _paragraph_style_id(paragraph: etree._Element) -> str | None:
    ppr = paragraph.find(PPR_TAG)
    if ppr is None:
        return None
    pstyle = ppr.find(PSTYLE_TAG)
    if pstyle is None:
        return None
    return pstyle.get(W_VAL_ATTR)


def _row_index_in_table(tbl: etree._Element, target_tr: etree._Element) -> int:
    idx = 0
    for tr in tbl.iter(TR_TAG):
        if _innermost_ancestor(tr, TBL_TAG) is not tbl:
            continue
        if tr is target_tr:
            return idx
        idx += 1
    raise ValueError("target_tr not found in tbl")


def _cell_indices(
    tr: etree._Element, target_tc: etree._Element
) -> tuple[int, int]:
    """Return ``(physical_cell_index, logical_column_index)`` for ``target_tc``."""
    physical = 0
    logical = _grid_before(tr)
    for tc in tr.iter(TC_TAG):
        if _innermost_ancestor(tc, TR_TAG) is not tr:
            continue
        if tc is target_tc:
            return physical, logical
        physical += 1
        logical += _grid_span(tc)
    raise ValueError("target_tc not found in tr")


def _paragraph_index_in_cell(tc: etree._Element, target_p: etree._Element) -> int:
    idx = 0
    for p in tc.iter(P_TAG):
        if _innermost_ancestor(p, TC_TAG) is not tc:
            continue
        if p is target_p:
            return idx
        idx += 1
    raise ValueError("target_p not found in tc")


def _table_nesting_depth(tbl: etree._Element) -> int:
    depth = 0
    parent = tbl.getparent()
    while parent is not None:
        if parent.tag == TBL_TAG:
            depth += 1
        parent = parent.getparent()
    return depth


def _table_context(
    paragraph: etree._Element, table_indices: dict[str, int]
) -> TableContext | None:
    tc = _innermost_ancestor(paragraph, TC_TAG)
    if tc is None:
        return None
    tr = _innermost_ancestor(tc, TR_TAG)
    if tr is None:
        return None
    tbl = _innermost_ancestor(tr, TBL_TAG)
    if tbl is None:
        return None
    physical, logical = _cell_indices(tr, tc)
    tree = tbl.getroottree()
    tbl_key = tree.getpath(tbl)
    return TableContext(
        table_index=table_indices[tbl_key],
        row_index=_row_index_in_table(tbl, tr),
        physical_cell_index=physical,
        logical_column_index=logical,
        grid_span=_grid_span(tc),
        grid_before=_grid_before(tr),
        vertical_merge=_vertical_merge(tc),
        nesting_depth=_table_nesting_depth(tbl),
        paragraph_index_in_cell=_paragraph_index_in_cell(tc, paragraph),
    )


def _build_table_indices(root: etree._Element) -> dict[str, int]:
    tree = root.getroottree()
    return {tree.getpath(tbl): idx for idx, tbl in enumerate(root.iter(TBL_TAG))}


def iter_editable_paragraphs(
    root: etree._Element,
) -> Iterator[tuple[etree._Element, ParagraphLocation, str | None, str | None]]:
    """Yield ``(paragraph, location, style_id, w14_para_id)`` in document order.

    Paragraphs beneath an SDT are skipped but still advance the global counter,
    so their structural contribution to ordinals and table positions is preserved.
    """
    table_indices = _build_table_indices(root)
    ordinal = 0
    for paragraph in root.iter(P_TAG):
        ordinal += 1
        if is_sdt_descendant(paragraph):
            continue
        location = ParagraphLocation(
            part="word/document.xml",
            original_index=ordinal,
            global_ordinal=ordinal,
            table_context=_table_context(paragraph, table_indices),
        )
        yield (
            paragraph,
            location,
            _paragraph_style_id(paragraph),
            paragraph.get(W14_PARA_ID_ATTR),
        )


def _under_deleted(elem: etree._Element, stop: etree._Element) -> bool:
    parent = elem.getparent()
    while parent is not None and parent is not stop:
        if parent.tag == DEL_TAG:
            return True
        parent = parent.getparent()
    return False


def visible_text_plain(paragraph: etree._Element) -> str:
    """Concatenate visible text of ``paragraph`` without reserved markers.

    Includes descendants of ``w:ins``, excludes descendants of ``w:del``,
    and projects tab/break nodes to ``\\t``, ``\\n``, ``\\r``.
    Phase 3 introduces ``visible_text_marked`` for TextMap-backed editing.
    """
    parts: list[str] = []
    for elem in paragraph.iter():
        tag = elem.tag
        if tag not in (T_TAG, TAB_TAG, BR_TAG, CR_TAG):
            continue
        if _under_deleted(elem, paragraph):
            continue
        if tag == T_TAG:
            if elem.text:
                parts.append(elem.text)
        elif tag == TAB_TAG:
            parts.append("\t")
        elif tag == BR_TAG:
            parts.append("\n")
        elif tag == CR_TAG:
            parts.append("\r")
    return "".join(parts)


def serialize_paragraph(paragraph: etree._Element) -> str:
    """Serialize ``paragraph`` including its start/end tags with no declaration."""
    raw = etree.tostring(paragraph, encoding="unicode", with_tail=False)
    return raw


__all__ = [
    "BR_TAG",
    "CR_TAG",
    "DELTEXT_TAG",
    "DEL_TAG",
    "GRID_BEFORE_TAG",
    "GRID_SPAN_TAG",
    "INS_TAG",
    "NSMAP",
    "PPR_TAG",
    "PSTYLE_TAG",
    "P_TAG",
    "R_TAG",
    "SDT_CONTENT_TAG",
    "SDT_TAG",
    "TAB_TAG",
    "TBL_TAG",
    "TCPR_TAG",
    "TC_TAG",
    "TRPR_TAG",
    "TR_TAG",
    "T_TAG",
    "VMERGE_TAG",
    "W14_NS",
    "W14_PARA_ID_ATTR",
    "W_NS",
    "W_VAL_ATTR",
    "XML_NS",
    "build_secure_parser",
    "is_sdt_descendant",
    "iter_editable_paragraphs",
    "qn",
    "serialize_paragraph",
    "visible_text_plain",
]
