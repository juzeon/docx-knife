"""Tests for :mod:`docx_knife.save` and ``Document.save``."""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

import pytest

from docx_knife import Document, SourceChangedError, ValidationError

from . import _fixtures

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _zip_entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _rewrite_source_bytes(path: Path, extra_entry: tuple[str, bytes]) -> None:
    """Externally edit ``path`` by appending an extra ZIP entry."""
    entries = _zip_entries(path)
    entries[extra_entry[0]] = extra_entry[1]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)


# ---------------------------------------------------------------------------
# 7.1 source drift
# ---------------------------------------------------------------------------


def test_save_detects_source_drift(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"
    with Document.open(src) as doc:
        _rewrite_source_bytes(src, ("extra.txt", b"tampered"))
        with pytest.raises(SourceChangedError) as excinfo:
            doc.save(dest)
        assert excinfo.value.source_path == str(src)
    assert not dest.exists()


def test_save_detects_source_deletion(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"
    with Document.open(src) as doc:
        src.unlink()
        with pytest.raises(SourceChangedError):
            doc.save(dest)
    assert not dest.exists()


# ---------------------------------------------------------------------------
# 7.1 same-path save (fingerprint refresh)
# ---------------------------------------------------------------------------


def test_save_same_path_updates_fingerprint(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    with Document.open(src) as doc:
        result = doc.save(src)
        assert result.output_path == str(src.resolve())
        # Backup captures the pre-save destination content.
        assert result.backup_path == str((tmp_path / "src.docx.bak").resolve())
        # A second save must not falsely detect drift.
        second = doc.save(src)
        assert second.output_path == str(src.resolve())

    with Document.open(src) as doc2:
        doc2.save(tmp_path / "out.docx")


# ---------------------------------------------------------------------------
# 7.2 / 7.4 new destination
# ---------------------------------------------------------------------------


def test_save_to_new_destination_has_no_backup(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "new.docx"
    with Document.open(src) as doc:
        result = doc.save(dest)
    assert result.backup_path is None
    assert result.output_path == str(dest.resolve())
    assert dest.is_file()
    with zipfile.ZipFile(dest, "r") as zf:
        assert "word/document.xml" in zf.namelist()
        # Reparses as XML.
        zf.read("word/document.xml").decode("utf-8")


# ---------------------------------------------------------------------------
# 7.2 repeated saves & backup replacement
# ---------------------------------------------------------------------------


def test_repeated_saves_keep_only_previous(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"

    with Document.open(src) as doc:
        first = doc.save(dest)
        first_bytes = dest.read_bytes()
        assert first.backup_path is None

    # Second document instance -> second save produces a .bak equal to first save's output.
    with Document.open(src) as doc2:
        second = doc2.save(dest)
        assert second.backup_path == str((tmp_path / "dest.docx.bak").resolve())
        bak_bytes = (tmp_path / "dest.docx.bak").read_bytes()
        assert bak_bytes == first_bytes


def test_existing_bak_is_replaced(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"
    bak = tmp_path / "dest.docx.bak"
    bak.write_bytes(b"OLD BACKUP CONTENT")

    dest.write_bytes(b"placeholder destination")
    with Document.open(src) as doc:
        doc.save(dest)
    assert bak.read_bytes() == b"placeholder destination"


# ---------------------------------------------------------------------------
# 7.4 injected failures
# ---------------------------------------------------------------------------


def test_write_failure_preserves_destination_and_backup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"
    original_dest_bytes = b"original destination bytes"
    dest.write_bytes(original_dest_bytes)

    from docx_knife import save as save_mod

    real_replace = os.replace

    def fake_replace(src_arg: object, dst_arg: object) -> None:
        # Fail only when replacing INTO the destination file.
        if Path(os.fspath(dst_arg)) == dest:
            raise OSError("simulated rename failure")
        real_replace(src_arg, dst_arg)

    monkeypatch.setattr(save_mod.os, "replace", fake_replace)

    with (
        Document.open(src) as doc,
        pytest.raises(OSError, match="simulated rename failure"),
    ):
        doc.save(dest)

    # Destination is untouched.
    assert dest.read_bytes() == original_dest_bytes
    # Backup was written before the failing rename and must survive.
    bak = tmp_path / "dest.docx.bak"
    assert bak.exists()
    assert bak.read_bytes() == original_dest_bytes


def test_validation_error_when_serialized_xml_malformed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "dest.docx"
    dest.write_bytes(b"existing destination")

    from docx_knife import save as save_mod

    monkeypatch.setattr(save_mod, "_serialize_main_xml", lambda _tree: b"<not-xml")

    with (
        Document.open(src) as doc,
        pytest.raises(ValidationError) as excinfo,
    ):
        doc.save(dest)
    assert excinfo.value.stage == "save"
    assert excinfo.value.failed_check == "xml_parse"

    # Destination unchanged.
    assert dest.read_bytes() == b"existing destination"
    # No backup written because validation runs BEFORE backup.
    assert not (tmp_path / "dest.docx.bak").exists()


# ---------------------------------------------------------------------------
# ZIP entries unchanged / unsupported structures untouched
# ---------------------------------------------------------------------------


def test_non_main_entries_byte_identical(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    dest = tmp_path / "out.docx"
    with Document.open(src) as doc:
        doc.save(dest)

    src_entries = _zip_entries(src)
    dst_entries = _zip_entries(dest)
    assert set(src_entries) == set(dst_entries)
    for name in src_entries:
        if name == "word/document.xml":
            continue
        assert src_entries[name] == dst_entries[name], f"entry drifted: {name}"


def test_sdt_document_round_trips(tmp_path: Path) -> None:
    src = _fixtures.build_sdt(tmp_path / "sdt.docx")
    dest = tmp_path / "sdt-out.docx"
    with Document.open(src) as doc:
        doc.save(dest)

    dst_main = _zip_entries(dest)["word/document.xml"].decode("utf-8")
    # The SDT-wrapped paragraph must survive.
    assert "w:sdt" in dst_main
    assert "<w:t>B</w:t>" in dst_main
    assert "<w:t>A</w:t>" in dst_main
    assert "<w:t>C</w:t>" in dst_main


# ---------------------------------------------------------------------------
# Closed-document guard
# ---------------------------------------------------------------------------


def test_save_on_closed_document_raises(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    doc = Document.open(src)
    doc.close()
    with pytest.raises(RuntimeError, match="closed"):
        doc.save(tmp_path / "out.docx")


def test_source_path_property(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "src.docx")
    with Document.open(src) as doc:
        assert doc.source_path == src
