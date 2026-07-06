"""Safe DOCX persistence: backup, ZIP rebuild, atomic replace (Phase 7).

The save pipeline is intentionally sequential and fail-fast:

1. Detect source drift against the fingerprint captured at ``open``.
2. Serialize ``word/document.xml`` in memory (no pretty-print).
3. Rebuild a temporary DOCX in the document's private scratch directory,
   preserving every non-main ZIP entry byte-for-byte and preserving entry
   metadata / order.
4. Reopen and validate the temporary package (ZIP integrity + main-XML parse).
5. Only after validation succeeds: replace ``<output>.bak`` (if any) with the
   current destination content, then atomically rename the temporary package
   into place.

Any failure before the atomic rename leaves the original destination intact.
A failure between backup and rename preserves the completed ``.bak`` so a human
operator can recover manually.
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import shutil
import uuid
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from . import _ooxml
from ._models import SaveResult
from .errors import SourceChangedError, ValidationError

if TYPE_CHECKING:
    from .document import Document, _SourceFingerprint

MAIN_PART: str = "word/document.xml"


def save_document(document: Document, output_path: str | os.PathLike[str]) -> SaveResult:
    """Persist ``document`` to ``output_path`` following the safe-save pipeline."""
    source_path: Path = document._source_path
    fingerprint = document._source_fingerprint

    _check_source_drift(source_path, fingerprint)

    dest = Path(output_path)
    dest_parent = dest.parent
    if not dest_parent.exists():
        raise FileNotFoundError(f"destination directory does not exist: {dest_parent}")

    scratch = document._temp_dir / f"save-{uuid.uuid4().hex}.docx"
    _build_temp_package(source_path, scratch, document._tree)
    _validate_temp_package(scratch)

    backup_path: Path | None = None
    if dest.exists():
        backup_path = _write_backup(dest)

    # Stage into dest's parent so os.replace never crosses drive boundaries.
    staged = dest_parent / f".docx-knife-save-{uuid.uuid4().hex}.docx"
    try:
        shutil.copy2(scratch, staged)
        os.replace(staged, dest)
    except BaseException:
        with contextlib.suppress(OSError):
            staged.unlink()
        raise
    finally:
        with contextlib.suppress(OSError):
            scratch.unlink()

    resolved_dest = dest.resolve()
    if resolved_dest == source_path.resolve():
        document._source_fingerprint = _fingerprint_from_disk(source_path)

    return SaveResult(
        output_path=str(resolved_dest),
        backup_path=str(backup_path.resolve()) if backup_path is not None else None,
        warnings=(),
    )


# ---------------------------------------------------------------------------
# Drift detection
# ---------------------------------------------------------------------------


def _check_source_drift(source_path: Path, expected: _SourceFingerprint) -> None:
    if not source_path.is_file():
        raise SourceChangedError(source_path=str(source_path))
    current = _fingerprint_from_disk(source_path)
    if (
        current.sha256 != expected.sha256
        or current.size != expected.size
        or current.mtime_ns != expected.mtime_ns
    ):
        raise SourceChangedError(source_path=str(source_path))


def _fingerprint_from_disk(path: Path) -> _SourceFingerprint:
    from .document import _SourceFingerprint  # local import to avoid cycle

    data = path.read_bytes()
    stat = path.stat()
    digest = hashlib.sha256(data).hexdigest()
    return _SourceFingerprint(sha256=digest, size=stat.st_size, mtime_ns=stat.st_mtime_ns)


# ---------------------------------------------------------------------------
# ZIP rebuild
# ---------------------------------------------------------------------------


def _serialize_main_xml(tree: etree._ElementTree) -> bytes:
    return etree.tostring(
        tree,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _build_temp_package(source_path: Path, temp_output: Path, tree: etree._ElementTree) -> None:
    new_main = _serialize_main_xml(tree)
    with (
        zipfile.ZipFile(source_path, "r") as source_zip,
        zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zip_out,
    ):
        seen_main = False
        for entry in source_zip.infolist():
            if entry.filename == MAIN_PART:
                new_info = _clone_zipinfo(entry)
                zip_out.writestr(new_info, new_main)
                seen_main = True
            else:
                data = source_zip.read(entry.filename)
                zip_out.writestr(_clone_zipinfo(entry), data)
        if not seen_main:
            raise ValidationError(
                stage="save",
                checks=("source_has_main",),
                failed_check="source_has_main",
            )


def _clone_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """Produce a fresh ZipInfo that preserves the source entry's metadata.

    ``writestr`` recomputes ``file_size`` and ``CRC``, but we retain compression
    type, timestamps, external attributes, comments and extras so untouched
    entries remain indistinguishable from the source.
    """
    clone = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
    clone.compress_type = info.compress_type
    clone.external_attr = info.external_attr
    clone.internal_attr = info.internal_attr
    clone.create_system = info.create_system
    clone.create_version = info.create_version
    clone.extract_version = info.extract_version
    clone.flag_bits = info.flag_bits
    clone.comment = info.comment
    clone.extra = info.extra
    return clone


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def _validate_temp_package(temp_output: Path) -> None:
    try:
        with zipfile.ZipFile(temp_output, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise ValidationError(
                    stage="save",
                    checks=("zip_reopen", "zip_testzip", "xml_parse"),
                    failed_check="zip_testzip",
                )
            if MAIN_PART not in zf.namelist():
                raise ValidationError(
                    stage="save",
                    checks=("zip_reopen", "main_present", "xml_parse"),
                    failed_check="main_present",
                )
            main_bytes = zf.read(MAIN_PART)
    except zipfile.BadZipFile as exc:
        raise ValidationError(
            stage="save",
            checks=("zip_reopen", "xml_parse"),
            failed_check="zip_reopen",
        ) from exc

    parser = _ooxml.build_secure_parser()
    try:
        etree.fromstring(main_bytes, parser)
    except etree.XMLSyntaxError as exc:
        raise ValidationError(
            stage="save",
            checks=("zip_reopen", "xml_parse"),
            failed_check="xml_parse",
        ) from exc


# ---------------------------------------------------------------------------
# Backup
# ---------------------------------------------------------------------------


def _write_backup(dest: Path) -> Path:
    """Atomically materialize ``<dest>.bak`` from the current ``dest`` bytes."""
    backup_path = dest.with_name(dest.name + ".bak")
    tmp = dest.with_name(f".{dest.name}.bak.{uuid.uuid4().hex}")
    try:
        shutil.copy2(dest, tmp)
        os.replace(tmp, backup_path)
    except BaseException:
        if tmp.exists():
            with contextlib.suppress(OSError):
                tmp.unlink()
        raise
    return backup_path


__all__ = ["save_document"]
