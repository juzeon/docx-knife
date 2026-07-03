"""End-to-end real-DOCX scenarios (Phase 8.1).

Each test drives the public ``docx_knife.Document`` workflow (open, query, edit,
save, reopen) against the realistic contract-like fixture built by
:func:`tests._fixtures.build_contract`. Nothing here mocks lxml or zipfile.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_knife import (
    BatchOperationError,
    Document,
    EditOperation,
    SourceChangedError,
)

from . import _fixtures

MAIN = "word/document.xml"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _zip_entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _main_xml_bytes(path: Path) -> bytes:
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(MAIN)


def _canonical(doc: Document, paragraph_id: str) -> bytes:
    return etree.tostring(doc._manifest.resolve(paragraph_id))


def _texts(doc: Document) -> list[str]:
    return [p.text or "" for p in doc.list_paragraphs(max_chars=0).paragraphs]


# ---------------------------------------------------------------------------
# Scenario 1 — Full round trip: open, query, mixed edits, save, reopen
# ---------------------------------------------------------------------------


def test_full_round_trip_open_query_edit_save_reopen(tmp_path: Path) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")
    out = tmp_path / "contract-edited.docx"

    with Document.open(src) as doc:
        listing = doc.list_paragraphs(max_chars=0)
        pre_count = doc.paragraph_count()
        assert pre_count == 30
        assert listing.pagination.total == 30
        ids = [p.id for p in listing.paragraphs]

        # Sanity queries before edits.
        assert doc.count_matches("target") >= 6
        first_target = doc.find_text("target", occurrence=0)
        assert first_target is not None
        # Discover a "target" paragraph through grep_paragraphs too.
        grep_hits = doc.grep_paragraphs("target")
        assert grep_hits.matches
        bracket_id = grep_hits.matches[-1].paragraph.id

        # Mixed batch exercising every top-level edit shape:
        #   * insert_para_before on the very first paragraph.
        #   * insert_para_after a Heading2 with two body paragraphs.
        #   * replace_text on a body paragraph containing "target".
        #   * insert_text_before / insert_text_after wrapping "target" with [].
        #   * replace_para on the hyperlink-bearing table paragraph (produces
        #     a "hyperlink" warning because wholesale replacement erases it).
        #   * delete_para on a plain trailing paragraph.
        head_anchor = ids[0]  # title paragraph
        after_anchor = ids[1]  # "第一部分 定义" (Heading2)
        text_anchor = ids[9]  # 第 1 条：责任条款 1 ... target ...
        table_anchor = ids[20]  # See Appendix A ... (hyperlink cell)
        del_anchor = ids[27]  # 附则条款 3.

        result = doc.batch_edit([
            EditOperation.insert_para_before(
                op_id="op-before",
                target_id=head_anchor,
                items=["前言："],
            ),
            EditOperation.insert_para_after(
                op_id="op-insert",
                target_id=after_anchor,
                items=["新增段落 A", "新增段落 B"],
            ),
            EditOperation.replace_text(
                op_id="op-replace-text",
                paragraph_id=text_anchor,
                find="target",
                replacement="TARGET",
            ),
            EditOperation.insert_text_before(
                op_id="op-itb",
                paragraph_id=bracket_id,
                find="target",
                text="[",
            ),
            EditOperation.insert_text_after(
                op_id="op-ita",
                paragraph_id=bracket_id,
                find="target",
                text="]",
            ),
            EditOperation.replace_para(
                op_id="op-replace-para",
                target_id=table_anchor,
                items=["Cell replacement — plain body."],
            ),
            EditOperation.delete_para(op_id="op-delete", target_ids=[del_anchor]),
        ])

        # Warning surfaces on both the OperationResult and the change log.
        rp_result = next(r for r in result.results if r.op_id == "op-replace-para")
        assert "hyperlink" in rp_result.warnings
        rp_log = next(e for e in doc.change_log() if e.get("op_id") == "op-replace-para")
        assert "hyperlink" in rp_log["warnings"]

        # Delta: +1 (before) + 2 (after) + 0 (replace 1->1) - 1 (delete) = +2.
        assert doc.paragraph_count() == pre_count + 2

        # Save to a fresh destination -> no backup.
        save_result = doc.save(out)
        assert save_result.backup_path is None
        assert save_result.output_path == str(out.resolve())

    # Reopen the saved file and verify semantics through the public API.
    with Document.open(out) as reopened:
        assert reopened.paragraph_count() == pre_count + 2
        texts = _texts(reopened)

        # Inserts are present and ordered.
        assert texts[0] == "前言："
        assert "新增段落 A" in texts and "新增段落 B" in texts
        assert texts.index("新增段落 A") < texts.index("新增段落 B")
        # replace_text edited paragraph is visible with TARGET.
        assert any("TARGET" in t and "责任条款 1" in t for t in texts)
        # insert_text_before/after wrapped "target" in [].
        assert any("[target]" in t for t in texts)
        # The table cell has been replaced with the plain body.
        assert "Cell replacement — plain body." in texts
        # The deleted paragraph is gone.
        assert "附则条款 3." not in texts
        # Ins content still visible; del content still hidden in visible mode.
        visible_all = reopened.get_visible_text()
        assert "INSERTED" in visible_all
        assert "SHOULD_BE_HIDDEN" not in visible_all

    # The SDT-wrapped paragraph survives as raw XML on disk, but was never
    # queryable through the public API on either doc.
    saved_main = _main_xml_bytes(out).decode("utf-8")
    src_main = _main_xml_bytes(src).decode("utf-8")
    assert "<w:sdt>" in saved_main and "<w:sdt>" in src_main
    assert "SDT-guarded paragraph body." in saved_main


# ---------------------------------------------------------------------------
# Scenario 2 — Rollback fidelity: byte-identical raw XML + one audit entry
# ---------------------------------------------------------------------------


def test_rollback_leaves_raw_visible_text_byte_identical(tmp_path: Path) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")
    out = tmp_path / "rolled-back.docx"

    with Document.open(src) as doc:
        raw_before = doc.get_visible_text(raw=True)
        ids = [p.id for p in doc.list_paragraphs().paragraphs]

        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.replace_text(
                    op_id="op-ok",
                    paragraph_id=ids[9],
                    find="target",
                    replacement="X",
                ),
                # Second op fails: the literal is not in that paragraph.
                EditOperation.replace_text(
                    op_id="op-bad",
                    paragraph_id=ids[10],
                    find="THIS_LITERAL_IS_ABSENT",
                    replacement="Y",
                ),
            ])

        # DOM restored exactly to the pre-batch bytes.
        assert doc.get_visible_text(raw=True) == raw_before

        # Exactly one failed-batch audit entry.
        failed = [
            entry
            for entry in doc.change_log()
            if entry.get("batch") is True and entry.get("status") == "failed"
        ]
        assert len(failed) == 1
        assert failed[0]["rolled_back"] is True

        # Save with no other edits still succeeds; reopened doc equals source.
        doc.save(out)

    with Document.open(src) as reference, Document.open(out) as reopened:
        assert reopened.get_visible_text() == reference.get_visible_text()


# ---------------------------------------------------------------------------
# Scenario 3 — Backup behavior across three consecutive saves
# ---------------------------------------------------------------------------


def test_backup_rolls_forward_across_three_saves(tmp_path: Path) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")
    out = tmp_path / "dest.docx"
    bak = out.with_name(out.name + ".bak")

    # First save: no existing dest → no backup.
    with Document.open(src) as doc:
        first = doc.save(out)
    assert first.backup_path is None
    first_bytes = out.read_bytes()

    # Second save: dest exists → .bak is the FIRST save's output.
    with Document.open(src) as doc:
        second = doc.save(out)
    assert second.backup_path == str(bak.resolve())
    assert bak.read_bytes() == first_bytes
    second_bytes = out.read_bytes()

    # Third save: .bak is now the SECOND save's output.
    with Document.open(src) as doc:
        third = doc.save(out)
    assert third.backup_path == str(bak.resolve())
    assert bak.read_bytes() == second_bytes


# ---------------------------------------------------------------------------
# Scenario 4 — Structural fidelity: non-main entries and untouched paragraphs
# ---------------------------------------------------------------------------


def test_non_main_entries_and_untouched_paragraphs_survive_save(
    tmp_path: Path,
) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")
    out = tmp_path / "edited.docx"

    src_entries = _zip_entries(src)

    with Document.open(src) as doc:
        pre_xml_by_index: list[bytes] = [
            _canonical(doc, p.id) for p in doc.list_paragraphs().paragraphs
        ]
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        # Edits at pre-index 9 (replace_text), 20 (replace_para), 27 (delete).
        touched_pre_indices = {9, 20, 27}

        doc.batch_edit([
            EditOperation.replace_text(
                op_id="op-r",
                paragraph_id=ids[9],
                find="target",
                replacement="TARGET",
            ),
            EditOperation.replace_para(
                op_id="op-rp",
                target_id=ids[20],
                items=["plain replacement"],
            ),
            EditOperation.delete_para(op_id="op-d", target_ids=[ids[27]]),
        ])
        doc.save(out)

    # Non-main ZIP entries are byte-identical.
    out_entries = _zip_entries(out)
    assert set(src_entries) == set(out_entries)
    for name in src_entries:
        if name == MAIN:
            continue
        assert src_entries[name] == out_entries[name], f"drift on {name!r}"

    # Reopen and compare canonical XML of every UNTOUCHED paragraph. The
    # pre/post index mapping shifts after the delete at pre-index 27:
    #   pre_idx <  20  → post_idx = pre_idx
    #   pre_idx == 20  → replaced anchor (touched)
    #   pre_idx <  27  → post_idx = pre_idx
    #   pre_idx == 27  → deleted
    #   pre_idx >  27  → post_idx = pre_idx - 1
    with Document.open(out) as reopened:
        post = reopened.list_paragraphs().paragraphs
        assert len(post) == 29  # 30 - 1 delete - 1 replace_old + 1 replace_new.

        for post_idx, info in enumerate(post):
            if post_idx < 20:
                pre_idx = post_idx
            elif post_idx == 20:
                continue  # replaced anchor position
            elif post_idx < 27:
                pre_idx = post_idx
            else:
                pre_idx = post_idx + 1
            if pre_idx in touched_pre_indices:
                continue
            assert _canonical(reopened, info.id) == pre_xml_by_index[pre_idx], (
                f"paragraph drifted at post_idx={post_idx} (pre_idx={pre_idx})"
            )


# ---------------------------------------------------------------------------
# Scenario 5 — Mid-batch failure via wrong-root raw fragment
# ---------------------------------------------------------------------------


def test_raw_fragment_with_wrong_root_rolls_back_batch(tmp_path: Path) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")

    with Document.open(src) as doc:
        ids = [p.id for p in doc.list_paragraphs().paragraphs]
        pre_texts = _texts(doc)
        pre_count = doc.paragraph_count()
        raw_before = doc.get_visible_text(raw=True)

        # Fragment root is <w:tbl>, not <w:p>; parse succeeds but the paragraph-
        # ops layer rejects it.
        bad_fragment = (
            '<w:tbl xmlns:w='
            '"http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:tr><w:tc><w:p><w:r><w:t>oops</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>"
        )
        with pytest.raises(BatchOperationError) as excinfo:
            doc.batch_edit([
                EditOperation.insert_para_after(
                    op_id="ok", target_id=ids[0], items=["fine"]
                ),
                EditOperation.insert_para_after(
                    op_id="bad",
                    target_id=ids[0],
                    items=[{"content_literal": bad_fragment}],
                    raw=True,
                ),
            ])
        assert excinfo.value.op_id == "bad"
        assert excinfo.value.rolled_back is True

        # DOM, paragraph_count and visible listing all restored.
        assert doc.paragraph_count() == pre_count
        assert _texts(doc) == pre_texts
        assert doc.get_visible_text(raw=True) == raw_before

        # Exactly one failed-batch audit event.
        failed = [
            entry
            for entry in doc.change_log()
            if entry.get("batch") is True and entry.get("status") == "failed"
        ]
        assert len(failed) == 1


# ---------------------------------------------------------------------------
# Scenario 6 — Source drift blocks save
# ---------------------------------------------------------------------------


def test_external_source_mutation_blocks_save(tmp_path: Path) -> None:
    src = _fixtures.build_contract(tmp_path / "contract.docx")
    dest = tmp_path / "shouldnt-exist.docx"
    src_bytes_before = src.read_bytes()

    with Document.open(src) as doc:
        # External tamper: rewrite the ZIP with an extra entry.
        entries = _zip_entries(src)
        entries["extra.txt"] = b"tampered"
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in entries.items():
                zf.writestr(name, data)
        src_bytes_after_tamper = src.read_bytes()
        assert src_bytes_after_tamper != src_bytes_before

        with pytest.raises(SourceChangedError):
            doc.save(dest)

    # Save did not touch either the destination or the tampered source.
    assert not dest.exists()
    assert src.read_bytes() == src_bytes_after_tamper
