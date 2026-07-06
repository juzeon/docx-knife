"""Tests for Document.list_sections and Document.get_section."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_knife import Document, ParagraphNotFoundError, SectionInfo

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


class TestListSections:
    def test_multiple_headings(self, tmp_path: Path) -> None:
        body = (
            _p("Intro", style="Heading1")
            + _p("Body 1")
            + _p("Body 2")
            + _p("Chapter 2", style="Heading1")
            + _p("Body 3")
        )
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            sections = doc.list_sections()
            assert len(sections) == 2
            assert sections[0].heading_text == "Intro"
            assert sections[0].level == 1
            assert len(sections[0].body_ids) == 2
            assert len(sections[0].all_ids) == 3
            assert sections[1].heading_text == "Chapter 2"
            assert len(sections[1].body_ids) == 1

    def test_level_filtering(self, tmp_path: Path) -> None:
        body = (
            _p("H1", style="Heading1")
            + _p("body")
            + _p("H2", style="Heading2")
            + _p("body under h2")
            + _p("Another H1", style="Heading1")
        )
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            all_sections = doc.list_sections()
            assert len(all_sections) == 3
            level1 = doc.list_sections(level=1)
            assert len(level1) == 2
            assert all(s.level == 1 for s in level1)
            level2 = doc.list_sections(level=2)
            assert len(level2) == 1
            assert level2[0].heading_text == "H2"

    def test_empty_document(self, tmp_path: Path) -> None:
        body = _p("Just a paragraph.") + _p("Another one.")
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            sections = doc.list_sections()
            assert sections == []

    def test_consecutive_headings(self, tmp_path: Path) -> None:
        body = (
            _p("First", style="Heading1")
            + _p("Second", style="Heading1")
            + _p("Body after second")
        )
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            sections = doc.list_sections()
            assert len(sections) == 2
            assert sections[0].body_ids == ()
            assert len(sections[1].body_ids) == 1

    def test_nested_headings_in_parent(self, tmp_path: Path) -> None:
        body = (
            _p("Parent", style="Heading1")
            + _p("body1")
            + _p("Child", style="Heading2")
            + _p("child body")
            + _p("Next Parent", style="Heading1")
        )
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            sections = doc.list_sections()
            h1_sections = [s for s in sections if s.level == 1]
            assert len(h1_sections) == 2
            parent = h1_sections[0]
            assert len(parent.body_ids) == 3


class TestGetSection:
    def test_valid_heading(self, tmp_path: Path) -> None:
        body = (
            _p("Title", style="Heading1")
            + _p("content")
            + _p("Sub", style="Heading2")
        )
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            paras = doc.list_paragraphs(max_chars=0).paragraphs
            heading_id = paras[0].id
            section = doc.get_section(heading_id)
            assert isinstance(section, SectionInfo)
            assert section.heading_text == "Title"
            assert section.level == 1

    def test_non_heading_raises(self, tmp_path: Path) -> None:
        body = _p("Title", style="Heading1") + _p("body text")
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            paras = doc.list_paragraphs(max_chars=0).paragraphs
            body_id = paras[1].id
            with pytest.raises(ValueError, match="not a heading"):
                doc.get_section(body_id)

    def test_invalid_id_raises(self, tmp_path: Path) -> None:
        body = _p("Title", style="Heading1")
        _write_docx(tmp_path / "t.docx", _wrap_document(body))
        with Document.open(tmp_path / "t.docx") as doc:
            with pytest.raises(ParagraphNotFoundError):
                doc.get_section("p_nonexistent")
