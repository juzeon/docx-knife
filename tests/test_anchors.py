"""Tests for the anchor manifest ID lifecycle."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_knife import Document, ParagraphNotFoundError

from . import _fixtures


def test_ids_are_unique_across_duplicate_paraid(tmp_path: Path) -> None:
    src = _fixtures.build_duplicate_paraid(tmp_path / "dupe.docx")
    with Document.open(src) as doc:
        listing = doc.list_paragraphs(max_chars=0)
        ids = [item.id for item in listing.paragraphs]
        assert ids == ["p_000001", "p_000002", "p_000003"]
        # The duplicates only survive as diagnostic metadata.
        paraids = [item.w14_para_id for item in listing.paragraphs]
        assert paraids == ["12345678", "12345678", None]


def test_invalidated_id_becomes_unresolvable(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        target = doc.list_paragraphs().paragraphs[0].id
        node = doc._manifest.resolve(target)
        parent = node.getparent()
        assert parent is not None
        parent.remove(node)
        doc._manifest.invalidate(target)
        doc._invalidate_index()
        with pytest.raises(ParagraphNotFoundError):
            doc._manifest.resolve(target)
        with pytest.raises(ParagraphNotFoundError):
            doc.get_paragraph(target)


def test_allocator_never_reuses_invalidated_id(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        manifest = doc._manifest
        first_id = manifest.ordered_ids()[0]
        manifest.invalidate(first_id)
        allocated = manifest.allocate()
        assert allocated != first_id
        # Existing IDs already went up to the count of paragraphs; the next allocation
        # must be strictly greater than the previous high-water mark.
        assert allocated == "p_000004"


def test_detached_node_reports_missing(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        target = doc.list_paragraphs().paragraphs[0].id
        node = doc._manifest.resolve(target)
        node.getparent().remove(node)  # type: ignore[union-attr]
        # Manifest still has the ID but the node is no longer reachable.
        with pytest.raises(ParagraphNotFoundError):
            doc._manifest.resolve(target)


def test_snapshot_restore_round_trips(tmp_path: Path) -> None:
    src = _fixtures.build_simple(tmp_path / "simple.docx")
    with Document.open(src) as doc:
        manifest = doc._manifest
        snapshot = manifest.snapshot()
        first_id = manifest.ordered_ids()[0]
        manifest.invalidate(first_id)
        _ = manifest.allocate()
        manifest.restore(snapshot)
        assert manifest.ordered_ids() == snapshot.order
        assert manifest.resolve(first_id) is not None
