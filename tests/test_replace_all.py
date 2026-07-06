"""Tests for Document.replace_all convenience method."""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx_knife import Document

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


def _p(text: str, *, style: str | None = None) -> str:
    ppr = ""
    if style is not None:
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/document.xml", document_xml)


def _texts(doc: Document) -> list[str]:
    return [p.text or "" for p in doc.list_paragraphs(max_chars=0).paragraphs]


def test_literal_replace_across_multiple_paragraphs(tmp_path: Path) -> None:
    body = _p("Old Corp is great.") + _p("Visit Old Corp today.") + _p("No match here.")
    _write_docx(tmp_path / "t.docx", _wrap_document(body))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all("Old Corp", "New Corp")
        assert count == 2
        texts = _texts(doc)
        assert texts[0] == "New Corp is great."
        assert texts[1] == "Visit New Corp today."
        assert texts[2] == "No match here."


def test_regex_replace(tmp_path: Path) -> None:
    body = _p("第1条 概述") + _p("第2条 定义") + _p("附录A")
    _write_docx(tmp_path / "t.docx", _wrap_document(body))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all(r"第\d+条", "第X条", regex=True)
        assert count == 2
        texts = _texts(doc)
        assert texts[0] == "第X条 概述"
        assert texts[1] == "第X条 定义"
        assert texts[2] == "附录A"


def test_zero_matches_returns_zero(tmp_path: Path) -> None:
    body = _p("Hello world.") + _p("Goodbye world.")
    _write_docx(tmp_path / "t.docx", _wrap_document(body))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all("nonexistent", "replacement")
        assert count == 0
        texts = _texts(doc)
        assert texts[0] == "Hello world."
        assert texts[1] == "Goodbye world."


def test_cross_run_match(tmp_path: Path) -> None:
    body_xml = (
        "<w:p>"
        "<w:r><w:t>hello </w:t></w:r>"
        "<w:r><w:t>world</w:t></w:r>"
        "</w:p>"
    )
    _write_docx(tmp_path / "t.docx", _wrap_document(body_xml))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all("hello world", "hi earth")
        assert count == 1
        texts = _texts(doc)
        assert texts[0] == "hi earth"


def test_formatting_preserved(tmp_path: Path) -> None:
    body_xml = (
        "<w:p>"
        '<w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>'
        "</w:p>"
    )
    _write_docx(tmp_path / "t.docx", _wrap_document(body_xml))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all("bold text", "new bold")
        assert count == 1
        xml = doc.get_paragraph(
            doc.list_paragraphs(max_chars=0).paragraphs[0].id, raw=True
        )
        assert "<w:b/>" in xml
        assert "new bold" in xml


def test_multiple_matches_in_one_paragraph(tmp_path: Path) -> None:
    body = _p("aaa bbb aaa bbb aaa")
    _write_docx(tmp_path / "t.docx", _wrap_document(body))
    with Document.open(tmp_path / "t.docx") as doc:
        count = doc.replace_all("aaa", "x")
        assert count == 3
        texts = _texts(doc)
        assert texts[0] == "x bbb x bbb x"
