"""Integration tests against real Word-authored DOCX files.

These fixtures live under ``tests/test_data/real/`` and were produced by
Microsoft Word / WPS, so unlike the hand-built fixtures in :mod:`._fixtures`
they exercise the full complexity of realistic OOXML: nested tables, headers,
footers, styles, numbering, custom XML parts, run-property splits, table-cell
paragraphs, and so on.

We drive every real fixture through the public ``docx_knife.Document`` API:
open → query → edit → save → reopen. The assertions are deliberately
schema-agnostic (counts, visible-text containment, non-main entry equality)
so they stay stable if Word authors ever touch these files.
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
)

MAIN = "word/document.xml"
REAL_DIR = Path(__file__).parent / "test_data" / "real"
REAL_FIXTURES = [
    "product-summary.docx",
    "contract.docx",
    "custody-agreement.docx",
    "prospectus.docx",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _zip_entries(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _visible_texts(doc: Document) -> list[str]:
    return [p.text or "" for p in doc.list_paragraphs(max_chars=0).paragraphs]


def _copy_fixture(name: str, tmp_path: Path) -> Path:
    src = REAL_DIR / name
    if not src.exists():
        pytest.skip(f"real fixture missing: {name}")
    dst = tmp_path / name
    dst.write_bytes(src.read_bytes())
    return dst


# ---------------------------------------------------------------------------
# Open + query
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("fixture_name", REAL_FIXTURES)
def test_open_and_enumerate_all_paragraphs(fixture_name: str, tmp_path: Path) -> None:
    src = _copy_fixture(fixture_name, tmp_path)
    with Document.open(src) as doc:
        page = doc.list_paragraphs(max_chars=0)
        total = doc.paragraph_count()
        assert total == page.pagination.total
        assert total == len(page.paragraphs)
        # Every paragraph exposes a stable id and a resolvable canonical form.
        ids = {p.id for p in page.paragraphs}
        assert len(ids) == total
        for info in page.paragraphs[:20]:
            xml = etree.tostring(doc._manifest.resolve(info.id))
            assert xml.startswith(b"<")


# ---------------------------------------------------------------------------
# Save fidelity: no-op round trip
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("fixture_name", REAL_FIXTURES)
def test_noop_save_preserves_content_and_non_main_entries(
    fixture_name: str, tmp_path: Path
) -> None:
    src = _copy_fixture(fixture_name, tmp_path)
    out = tmp_path / f"noop-{fixture_name}"

    src_entries = _zip_entries(src)
    with Document.open(src) as doc:
        pre_count = doc.paragraph_count()
        pre_visible = doc.get_visible_text()
        pre_raw = doc.get_visible_text(raw=True)
        doc.save(out)

    out_entries = _zip_entries(out)
    assert set(src_entries) == set(out_entries)
    for name, blob in src_entries.items():
        if name == MAIN:
            continue
        assert out_entries[name] == blob, f"non-main entry drifted: {name!r}"

    with Document.open(out) as reopened:
        assert reopened.paragraph_count() == pre_count
        assert reopened.get_visible_text() == pre_visible
        assert reopened.get_visible_text(raw=True) == pre_raw


# ---------------------------------------------------------------------------
# Product summary — targeted edit exercise
# ---------------------------------------------------------------------------


def test_product_summary_batch_edit_round_trip(tmp_path: Path) -> None:
    """Drive a realistic mixed batch against the ETF product summary.

    Chosen edits:
      * ``replace_text`` on the ``编制日期`` line to inject a concrete date.
      * ``insert_text_after`` on the ``送出日期`` line to append a note.
      * ``insert_para_after`` right after the ``一、产品概况`` heading.
      * ``replace_para`` on the ``基金简称`` table cell (a paragraph inside
        ``<w:tc>``), which produces no warnings (no hyperlink).
      * ``delete_para`` on a plain trailing metadata paragraph.
    """
    src = _copy_fixture("product-summary.docx", tmp_path)
    out = tmp_path / "edited.docx"

    with Document.open(src) as doc:
        pre_count = doc.paragraph_count()

        edit_date_id = doc.grep_paragraphs("编制日期").matches[0].paragraph.id
        send_date_id = doc.grep_paragraphs("送出日期").matches[0].paragraph.id
        heading_id = doc.grep_paragraphs("一、产品概况").matches[0].paragraph.id
        cell_id = doc.grep_paragraphs("基金简称").matches[0].paragraph.id

        # Pick a delete target that is: not inside a table cell (avoid tampering
        # with table structure), non-empty, and clearly late in the document.
        late_targets = []
        for p in doc.list_paragraphs().paragraphs:
            text = p.text or ""
            if not text.strip():
                continue
            parent = doc._manifest.resolve(p.id).getparent()
            if parent is None or not parent.tag.endswith("}body"):
                continue
            late_targets.append(p)
        assert late_targets, "expected at least one body-level paragraph"
        del_target = late_targets[-1]
        del_target_text = del_target.text or ""

        result = doc.batch_edit([
            EditOperation.replace_text(
                op_id="op-r",
                paragraph_id=edit_date_id,
                find="2024年X月X日",
                replacement="2026年7月3日",
            ),
            EditOperation.insert_text_after(
                op_id="op-ita",
                paragraph_id=send_date_id,
                find="2024年X月X日",
                text="（预填）",
            ),
            EditOperation.insert_para_after(
                op_id="op-ipa",
                target_id=heading_id,
                items=["【补充说明】以下为示例条款。"],
            ),
            EditOperation.replace_para(
                op_id="op-rp",
                target_id=cell_id,
                items=["基金简称（新）"],
            ),
            EditOperation.delete_para(op_id="op-d", target_ids=[del_target.id]),
        ])

        assert [r.op_id for r in result.results] == [
            "op-r",
            "op-ita",
            "op-ipa",
            "op-rp",
            "op-d",
        ]
        # replace_para on a plain paragraph should not emit hyperlink warnings.
        rp = next(r for r in result.results if r.op_id == "op-rp")
        assert "hyperlink" not in rp.warnings

        # Delta: +1 insert - 1 delete = 0.
        assert doc.paragraph_count() == pre_count

        doc.save(out)

    with Document.open(out) as reopened:
        texts = _visible_texts(reopened)
        assert any("编制日期：2026年7月3日" in t for t in texts)
        assert any("送出日期：2024年X月X日（预填）" in t for t in texts)
        assert any("【补充说明】以下为示例条款。" in t for t in texts)
        assert "基金简称（新）" in texts
        assert del_target_text not in texts


# ---------------------------------------------------------------------------
# Contract — replace_text uniqueness handling on a heavily repeated token
# ---------------------------------------------------------------------------


def test_contract_replace_text_across_multiple_hits(tmp_path: Path) -> None:
    """The fund contract mentions ``招商银行`` several times. We replace them
    one occurrence at a time, verifying ``occurrence`` semantics on a real
    document with many run-property splits.
    """
    src = _copy_fixture("contract.docx", tmp_path)
    out = tmp_path / "contract-edited.docx"

    with Document.open(src) as doc:
        before = doc.count_matches("招商银行")
        assert before >= 2, "fixture expected to contain multiple hits"

        # Replace only the first occurrence.
        first_para_id = doc.grep_paragraphs("招商银行").matches[0].paragraph.id
        doc.batch_edit([
            EditOperation.replace_text(
                op_id="first",
                paragraph_id=first_para_id,
                find="招商银行",
                replacement="示例银行",
                occurrence=0,
            )
        ])
        after_first = doc.count_matches("招商银行")
        assert after_first == before - 1
        assert doc.count_matches("示例银行") >= 1

        doc.save(out)

    with Document.open(out) as reopened:
        # Confirm the delta persisted.
        assert reopened.count_matches("招商银行") == before - 1
        assert reopened.count_matches("示例银行") >= 1


# ---------------------------------------------------------------------------
# Prospectus — batch rollback fidelity on a large real document
# ---------------------------------------------------------------------------


def test_prospectus_rollback_leaves_document_untouched(tmp_path: Path) -> None:
    """Force a mid-batch failure on a 2000+ paragraph real document and
    confirm the DOM is byte-identical afterwards (raw visible text)."""
    src = _copy_fixture("prospectus.docx", tmp_path)

    with Document.open(src) as doc:
        raw_before = doc.get_visible_text(raw=True)
        pre_count = doc.paragraph_count()

        first_id = doc.list_paragraphs().paragraphs[0].id
        # First op is valid but arbitrary; the second op targets a literal
        # guaranteed not to exist inside the first paragraph.
        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.insert_para_after(
                    op_id="ok",
                    target_id=first_id,
                    items=["ROLL_BACK_ME"],
                ),
                EditOperation.replace_text(
                    op_id="bad",
                    paragraph_id=first_id,
                    find="__NOT_IN_THIS_DOC__",
                    replacement="X",
                ),
            ])

        assert doc.paragraph_count() == pre_count
        assert doc.get_visible_text(raw=True) == raw_before
        assert "ROLL_BACK_ME" not in doc.get_visible_text()

        failed = [
            entry
            for entry in doc.change_log()
            if entry.get("batch") is True and entry.get("status") == "failed"
        ]
        assert len(failed) == 1
        assert failed[0]["rolled_back"] is True
