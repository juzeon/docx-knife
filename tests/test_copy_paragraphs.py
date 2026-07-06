"""Tests for Document.copy_paragraphs_from."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_knife import Document, ParagraphNotFoundError

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


def _p(text: str, *, style: str | None = None, bold: bool = False) -> str:
    ppr = ""
    if style is not None:
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f"<w:p>{ppr}<w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>"


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/document.xml", document_xml)


def test_range_copy(tmp_path: Path) -> None:
    body = _p("Para 1") + _p("Para 2") + _p("Para 3") + _p("Para 4")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        start_id = paras[1].id
        end_id = paras[2].id
        result = target.copy_paragraphs_from(source, start_id, end_id)
        assert len(result) == 2
        assert "Para 2" in result[0]
        assert "Para 3" in result[1]


def test_single_paragraph(tmp_path: Path) -> None:
    body = _p("Only one") + _p("Another")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        pid = paras[0].id
        result = target.copy_paragraphs_from(source, pid, pid)
        assert len(result) == 1
        assert "Only one" in result[0]


def test_invalid_start_id(tmp_path: Path) -> None:
    body = _p("Hello")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        with pytest.raises(ParagraphNotFoundError):
            target.copy_paragraphs_from(source, "p_invalid", paras[0].id)


def test_invalid_end_id(tmp_path: Path) -> None:
    body = _p("Hello")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        with pytest.raises(ParagraphNotFoundError):
            target.copy_paragraphs_from(source, paras[0].id, "p_invalid")


def test_reversed_range_raises(tmp_path: Path) -> None:
    body = _p("First") + _p("Second") + _p("Third")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        with pytest.raises(ValueError, match="appears after"):
            target.copy_paragraphs_from(source, paras[2].id, paras[0].id)


def test_namespace_present(tmp_path: Path) -> None:
    body = _p("Content")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        result = target.copy_paragraphs_from(source, paras[0].id, paras[0].id)
        assert 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' in result[0]


def test_formatting_preserved(tmp_path: Path) -> None:
    body = _p("Bold text", bold=True) + _p("Italic", style="Heading1")
    src_path = tmp_path / "source.docx"
    tgt_path = tmp_path / "target.docx"
    _write_docx(src_path, _wrap_document(body))
    _write_docx(tgt_path, _wrap_document(_p("Target")))

    with Document.open(src_path) as source, Document.open(tgt_path) as target:
        paras = source.list_paragraphs(max_chars=0).paragraphs
        result = target.copy_paragraphs_from(source, paras[0].id, paras[1].id)
        assert len(result) == 2
        assert "<w:b/>" in result[0]
        assert "Heading1" in result[1]
