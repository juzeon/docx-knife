"""Programmatic DOCX fixture builders for Phase 2 tests.

We build the ZIP + XML by hand rather than depending on ``python-docx`` so the
fixtures stay explicit about the exact structures under test (SDT wrappers,
nested tables, duplicate ``w14:paraId`` etc.).
"""

from __future__ import annotations

import zipfile
from pathlib import Path

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="word/document.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
</Relationships>
"""


def _wrap_document(body_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{_W_NS}" xmlns:w14="{_W14_NS}">'
        f"<w:body>{body_xml}</w:body>"
        "</w:document>"
    )


def _p(text: str, *, style: str | None = None, w14_id: str | None = None) -> str:
    ppr = ""
    if style is not None:
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    attrs = ""
    if w14_id is not None:
        attrs = f' w14:paraId="{w14_id}"'
    return f"<w:p{attrs}>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/document.xml", document_xml)


def build_simple(path: Path) -> Path:
    body = _p("First paragraph.") + _p("A heading.", style="Heading1") + _p("Third paragraph.")
    _write_docx(path, _wrap_document(body))
    return path


def build_duplicate_paraid(path: Path) -> Path:
    body = (
        _p("Alpha.", w14_id="12345678")
        + _p("Beta.", w14_id="12345678")
        + _p("Gamma.")
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_sdt(path: Path) -> Path:
    """A body with plain paragraph A, an SDT-wrapped paragraph B, then plain C."""
    body = (
        _p("A")
        + "<w:sdt><w:sdtContent>" + _p("B") + "</w:sdtContent></w:sdt>"
        + _p("C")
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_nested_tables(path: Path) -> Path:
    """Outer 2x2 table where cell (0,0) contains an inner 1x2 table.

    Row 1 has ``gridSpan=2`` on its first cell and ``gridBefore=1`` on the row,
    so we can exercise logical-column arithmetic.
    """
    inner_table = (
        "<w:tbl>"
        "<w:tr>"
        "<w:tc>" + _p("inner-a") + "</w:tc>"
        "<w:tc>" + _p("inner-b") + "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    outer = (
        "<w:tbl>"
        "<w:tr>"
        "<w:tc>" + inner_table + _p("outer-0-0-tail") + "</w:tc>"
        "<w:tc>" + _p("outer-0-1") + "</w:tc>"
        "</w:tr>"
        "<w:tr>"
        '<w:trPr><w:gridBefore w:val="1"/></w:trPr>'
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' + _p("outer-1-span") + "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    body = _p("preamble") + outer + _p("epilogue")
    _write_docx(path, _wrap_document(body))
    return path


def build_raw_query(path: Path) -> Path:
    body = (
        "<w:p><w:r><w:t>违约</w:t></w:r>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>责任</w:t></w:r>"
        "<w:r><w:t>约定：三十日内履行。</w:t></w:r></w:p>"
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_abcd(path: Path) -> Path:
    """Four plain sibling paragraphs A, B, C, D — the workhorse ordering fixture."""
    body = _p("A") + _p("B") + _p("C") + _p("D")
    _write_docx(path, _wrap_document(body))
    return path


def build_formatted(path: Path) -> Path:
    """One paragraph with Heading1 style and a bold ordinary text run."""
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Anchor text.</w:t></w:r></w:p>"
        + _p("Second.")
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_hyperlink_para(path: Path) -> Path:
    """Paragraph containing a hyperlink plus a locally bold run."""
    body = (
        "<w:p>"
        '<w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:r><w:t>anchor</w:t></w:r>"
        "</w:hyperlink>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t> tail</w:t></w:r>"
        "</w:p>"
        + _p("After.")
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_table_paragraph(path: Path) -> Path:
    """Body with a leading paragraph and a nested single-cell table."""
    inner_table = (
        "<w:tbl>"
        "<w:tr>"
        "<w:tc>" + _p("cell-first") + "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    body = _p("preamble") + inner_table + _p("epilogue")
    _write_docx(path, _wrap_document(body))
    return path


__all__ = [
    "build_abcd",
    "build_duplicate_paraid",
    "build_formatted",
    "build_hyperlink_para",
    "build_nested_tables",
    "build_raw_query",
    "build_sdt",
    "build_simple",
    "build_table_paragraph",
]
