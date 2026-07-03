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
    body = _p("Alpha.", w14_id="12345678") + _p("Beta.", w14_id="12345678") + _p("Gamma.")
    _write_docx(path, _wrap_document(body))
    return path


def build_sdt(path: Path) -> Path:
    """A body with plain paragraph A, an SDT-wrapped paragraph B, then plain C."""
    body = _p("A") + "<w:sdt><w:sdtContent>" + _p("B") + "</w:sdtContent></w:sdt>" + _p("C")
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
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Anchor text.</w:t></w:r></w:p>" + _p("Second.")
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
        "</w:p>" + _p("After.")
    )
    _write_docx(path, _wrap_document(body))
    return path


def build_table_paragraph(path: Path) -> Path:
    """Body with a leading paragraph and a nested single-cell table."""
    inner_table = "<w:tbl><w:tr><w:tc>" + _p("cell-first") + "</w:tc></w:tr></w:tbl>"
    body = _p("preamble") + inner_table + _p("epilogue")
    _write_docx(path, _wrap_document(body))
    return path


_R_NS_DECL = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def _hyperlink_para() -> str:
    return (
        "<w:p>"
        f'<w:hyperlink r:id="rId42" {_R_NS_DECL}>'
        "<w:r><w:t>See Appendix A</w:t></w:r>"
        "</w:hyperlink>"
        "<w:r><w:t> for details.</w:t></w:r>"
        "</w:p>"
    )


def _tab_br_para() -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
        "<w:r><w:t>Left</w:t><w:tab/><w:t>Middle</w:t>"
        "<w:br/><w:t>NextLine</w:t></w:r>"
        "</w:p>"
    )


def _ins_para() -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
        "<w:r><w:t>Base text </w:t></w:r>"
        '<w:ins w:id="1" w:author="Alice" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:t>INSERTED</w:t></w:r>"
        "</w:ins>"
        "<w:r><w:t> tail.</w:t></w:r>"
        "</w:p>"
    )


def _del_para() -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
        "<w:r><w:t>Keep me </w:t></w:r>"
        '<w:del w:id="2" w:author="Bob" w:date="2024-01-02T00:00:00Z">'
        "<w:r><w:delText>SHOULD_BE_HIDDEN</w:delText></w:r>"
        "</w:del>"
        "<w:r><w:t> remainder.</w:t></w:r>"
        "</w:p>"
    )


def build_contract(path: Path) -> Path:
    """A contract-like DOCX exercising every structural feature the e2e tests need.

    Structure (in document order):

    * Heading1 title paragraph.
    * Two Heading2 section paragraphs.
    * A run of Normal body paragraphs (numbered clauses).
    * Two paragraphs sharing the same ``w14:paraId`` to prove the anchor
      manifest allocates distinct instance-local IDs regardless of Word's
      duplicate hints.
    * A tab/br paragraph.
    * An ``<w:ins>`` paragraph and a ``<w:del>`` paragraph.
    * A 2x2 nested table where cell (0,0) hosts a hyperlink paragraph.
    * An SDT block wrapping a single paragraph (must remain uneditable).
    * A tail of Normal body paragraphs, followed by a signature block.
    """
    # 3 heading + 12 clause + 2 duplicate-paraId + 1 tab/br + 1 ins + 1 del
    # = 20 non-table editable body paragraphs before the table.
    body_head: list[str] = []
    body_head.append(_p("契约合同 — Test Contract", style="Heading1"))
    body_head.append(_p("第一部分 定义", style="Heading2"))
    for i in range(1, 7):
        body_head.append(_p(f"第 {i} 条：定义条款 {i} 的正文内容。"))
    body_head.append(_p("第二部分 责任", style="Heading2"))
    for i in range(1, 7):
        body_head.append(_p(f"第 {i} 条：责任条款 {i} 的正文，包含关键字 target。"))

    body_head.append(_p("双 paraId 段落 alpha。", w14_id="ABCD1234"))
    body_head.append(_p("双 paraId 段落 beta。", w14_id="ABCD1234"))
    body_head.append(_tab_br_para())
    body_head.append(_ins_para())
    body_head.append(_del_para())

    table = (
        "<w:tbl>"
        "<w:tr>"
        "<w:tc>" + _hyperlink_para() + "</w:tc>"
        "<w:tc>" + _p("Cell (0,1) plain.") + "</w:tc>"
        "</w:tr>"
        "<w:tr>"
        "<w:tc>" + _p("Cell (1,0) plain.") + "</w:tc>"
        "<w:tc>" + _p("Cell (1,1) plain.") + "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )

    sdt_block = (
        "<w:sdt><w:sdtContent>" + _p("SDT-guarded paragraph body.") + "</w:sdtContent></w:sdt>"
    )

    body_tail: list[str] = []
    body_tail.append(_p("第三部分 附则", style="Heading2"))
    for i in range(1, 4):
        body_tail.append(_p(f"附则条款 {i}."))
    body_tail.append(_p("签署人：甲方", style="Normal"))
    body_tail.append(_p("签署人：乙方", style="Normal"))

    body = "".join(body_head) + table + sdt_block + "".join(body_tail)
    _write_docx(path, _wrap_document(body))
    return path


__all__ = [
    "build_abcd",
    "build_contract",
    "build_duplicate_paraid",
    "build_formatted",
    "build_hyperlink_para",
    "build_nested_tables",
    "build_raw_query",
    "build_sdt",
    "build_simple",
    "build_table_paragraph",
]
