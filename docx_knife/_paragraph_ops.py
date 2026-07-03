"""Whole-paragraph operations (Phase 4).

Helpers that construct new ``<w:p>`` elements from visible text or raw
fragments, and inspect paragraphs for structures that a wholesale replacement
would destroy. All names are module-private; the public surface is
:class:`~docx_knife.paragraph.Paragraph` and the ``Document`` methods that
delegate here.
"""

from __future__ import annotations

import copy
from collections.abc import Sequence

from lxml import etree

from . import _ooxml
from ._ooxml import qn
from .content import ResolvedItem
from .errors import InvalidContentError
from .textmap import restore_markers

_R_TAG = _ooxml.R_TAG
_T_TAG = _ooxml.T_TAG
_P_TAG = _ooxml.P_TAG
_PPR_TAG = _ooxml.PPR_TAG
_TAB_TAG = _ooxml.TAB_TAG
_BR_TAG = _ooxml.BR_TAG
_CR_TAG = _ooxml.CR_TAG
_INS_TAG = _ooxml.INS_TAG
_DEL_TAG = _ooxml.DEL_TAG
_RPR_TAG = qn("w:rPr")
_FLDCHAR_TAG = qn("w:fldChar")
_INSTR_TEXT_TAG = qn("w:instrText")
_HYPERLINK_TAG = qn("w:hyperlink")
_BOOKMARK_START_TAG = qn("w:bookmarkStart")
_BOOKMARK_END_TAG = qn("w:bookmarkEnd")
_PERM_START_TAG = qn("w:permStart")
_PERM_END_TAG = qn("w:permEnd")
_COMMENT_START_TAG = qn("w:commentRangeStart")
_COMMENT_END_TAG = qn("w:commentRangeEnd")
_TYPE_ATTR = qn("w:type")
_XML_SPACE_ATTR = f"{{{_ooxml.XML_NS}}}space"

_MARKER_TAGS: dict[str, str] = {
    "TAB": _TAB_TAG,
    "LINE_BREAK": _BR_TAG,
    "PAGE_BREAK": _BR_TAG,
    "COLUMN_BREAK": _BR_TAG,
    "CR": _CR_TAG,
}
_MARKER_ATTR: dict[str, str | None] = {
    "TAB": None,
    "LINE_BREAK": None,
    "PAGE_BREAK": "page",
    "COLUMN_BREAK": "column",
    "CR": None,
}

# Detection tag -> stable label.
_PROTECTED_ITER_TAGS: tuple[tuple[str, str], ...] = (
    (_HYPERLINK_TAG, "hyperlink"),
    (_FLDCHAR_TAG, "field"),
    (_INSTR_TEXT_TAG, "field"),
    (_BOOKMARK_START_TAG, "bookmark"),
    (_BOOKMARK_END_TAG, "bookmark"),
    (_PERM_START_TAG, "permission"),
    (_PERM_END_TAG, "permission"),
    (_COMMENT_START_TAG, "comment"),
    (_COMMENT_END_TAG, "comment"),
    (_INS_TAG, "revision_ins"),
    (_DEL_TAG, "revision_del"),
)


def expand_visible_items(items: Sequence[ResolvedItem]) -> list[str]:
    """Flatten resolved visible-mode items into an ordered list of paragraph texts.

    Raises :class:`InvalidContentError` if any item was resolved in raw mode.
    """
    paragraphs: list[str] = []
    for item in items:
        if item.raw:
            raise InvalidContentError(
                raw=True,
                reason="mixed content modes not allowed",
            )
        paragraphs.extend(item.paragraphs)
    return paragraphs


def _find_first_text_run_rpr(anchor: etree._Element) -> etree._Element | None:
    """Return a deep copy of the anchor's first ordinary text run's ``<w:rPr>``.

    An ordinary text run has at least one ``<w:t>`` child and no
    ``<w:fldChar>`` or ``<w:instrText>`` children.
    """
    for run in anchor.iterchildren(_R_TAG):
        has_text = False
        disqualified = False
        for child in run:
            tag = child.tag
            if tag == _T_TAG:
                has_text = True
            elif tag in (_FLDCHAR_TAG, _INSTR_TEXT_TAG):
                disqualified = True
                break
        if has_text and not disqualified:
            rpr = run.find(_RPR_TAG)
            return copy.deepcopy(rpr) if rpr is not None else None
    return None


def _append_run_with_rpr(
    parent: etree._Element, rpr_template: etree._Element | None
) -> etree._Element:
    run = etree.SubElement(parent, _R_TAG)
    if rpr_template is not None:
        run.append(copy.deepcopy(rpr_template))
    return run


def _append_text_child(run: etree._Element, text: str) -> None:
    if not text:
        return
    t = etree.SubElement(run, _T_TAG)
    if text[:1].isspace() or text[-1:].isspace():
        t.set(_XML_SPACE_ATTR, "preserve")
    t.text = text


def _append_marker_child(run: etree._Element, marker: str) -> None:
    tag = _MARKER_TAGS[marker]
    node = etree.SubElement(run, tag)
    attr_val = _MARKER_ATTR[marker]
    if attr_val is not None:
        node.set(_TYPE_ATTR, attr_val)


def build_new_paragraph(anchor: etree._Element, text: str) -> etree._Element:
    """Construct a new ``<w:p>`` inheriting formatting from ``anchor``.

    * ``<w:pPr>`` is deep-copied.
    * The first ordinary text run's ``<w:rPr>`` is deep-copied and reused for
      every emitted run.
    * Embedded single ``\\n`` characters project to ``<w:br/>``.
    * Reserved ``[[DOCX:...]]`` markers project to their real XML nodes.
    * Text with leading/trailing whitespace sets ``xml:space="preserve"``.

    Unknown markers raise :class:`InvalidContentError`.
    """
    new_p = etree.Element(_P_TAG)
    ppr = anchor.find(_PPR_TAG)
    if ppr is not None:
        new_p.append(copy.deepcopy(ppr))
    rpr_template = _find_first_text_run_rpr(anchor)

    projected = text.replace("\n", "[[DOCX:LINE_BREAK]]")
    segments = restore_markers(projected)
    for kind, payload in segments:
        run = _append_run_with_rpr(new_p, rpr_template)
        if kind == "text":
            _append_text_child(run, payload)
        else:
            _append_marker_child(run, payload)
    return new_p


def parse_raw_paragraphs(fragment: str) -> list[etree._Element]:
    """Parse a raw XML fragment containing one or more top-level ``<w:p>`` nodes.

    The fragment is wrapped in a namespace-safe root so bare ``<w:p>`` fragments
    parse correctly. All top-level nodes must be ``<w:p>``; anything else, or a
    malformed fragment, raises :class:`InvalidContentError`.
    """
    wrapper = (
        '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        f"{fragment}"
        "</root>"
    )
    parser = _ooxml.build_secure_parser()
    try:
        root = etree.fromstring(wrapper.encode("utf-8"), parser)
    except etree.XMLSyntaxError as exc:
        raise InvalidContentError(
            raw=True,
            reason=f"fragment is not well-formed XML: {exc.msg}",
        ) from exc
    children = list(root)
    if not children:
        raise InvalidContentError(
            raw=True,
            reason="raw fragment must contain at least one <w:p>",
        )
    for child in children:
        if child.tag != _P_TAG:
            raise InvalidContentError(
                raw=True,
                reason=f"raw fragment top-level element must be w:p: found {child.tag}",
            )
    for child in children:
        root.remove(child)
    return children


def detect_protected_structures(paragraph: etree._Element) -> tuple[str, ...]:
    """Return stable-ordered labels of structures a wholesale replace would erase."""
    found: list[str] = []
    seen: set[str] = set()

    def _add(label: str) -> None:
        if label not in seen:
            seen.add(label)
            found.append(label)

    for tag, label in _PROTECTED_ITER_TAGS:
        # Only need to know presence; break after first hit per label.
        for _ in paragraph.iter(tag):
            _add(label)
            break

    for run in paragraph.iter(_R_TAG):
        if run.find(_RPR_TAG) is not None:
            _add("run_styling")
            break

    return tuple(found)


__all__ = [
    "build_new_paragraph",
    "detect_protected_structures",
    "expand_visible_items",
    "parse_raw_paragraphs",
]
