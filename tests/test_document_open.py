"""Tests for ``Document.open`` and lifecycle behavior."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_knife import Document, DocumentNotFoundError, InvalidDocumentError

from . import _fixtures


def test_open_missing_file_raises(tmp_path: Path) -> None:
    missing = tmp_path / "nope.docx"
    with pytest.raises(DocumentNotFoundError) as excinfo:
        Document.open(missing)
    assert excinfo.value.path == str(missing)


def test_open_non_zip_raises_invalid(tmp_path: Path) -> None:
    fake = tmp_path / "not.docx"
    fake.write_text("this is not a zip", encoding="utf-8")
    with pytest.raises(InvalidDocumentError) as excinfo:
        Document.open(fake)
    assert "ZIP" in excinfo.value.reason or "zip" in excinfo.value.reason.lower()


def test_open_zip_without_main_document_raises(tmp_path: Path) -> None:
    junk = tmp_path / "no_main.docx"
    with zipfile.ZipFile(junk, "w") as zf:
        zf.writestr("only.txt", "hello")
    with pytest.raises(InvalidDocumentError) as excinfo:
        Document.open(junk)
    assert "word/document.xml" in excinfo.value.reason


def test_open_zip_with_unparseable_main_raises(tmp_path: Path) -> None:
    bad = tmp_path / "bad.docx"
    with zipfile.ZipFile(bad, "w") as zf:
        zf.writestr("word/document.xml", "<not-xml")
    with pytest.raises(InvalidDocumentError) as excinfo:
        Document.open(bad)
    assert "cannot parse" in excinfo.value.reason


def test_context_manager_cleans_temp_dir(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        temp_dir = doc._temp_dir
        assert temp_dir.exists()
    assert not temp_dir.exists()


def test_source_fingerprint_captured(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        fp = doc._fingerprint
        assert len(fp.sha256) == 64
        assert fp.size == src.stat().st_size
        assert fp.mtime_ns == src.stat().st_mtime_ns
