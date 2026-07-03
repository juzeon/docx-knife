"""Atomic batch execution, conflict detection, and rollback (Phase 6).

The public surface is :class:`Document.batch_edit`; ``BatchExecutor`` is a
private orchestrator that owns prevalidation, deterministic ordering,
snapshotting, execution, and precommit checks. Any failure at any stage
restores the pre-batch DOM, anchor manifest, and change-log tail exactly and
raises :class:`BatchOperationError` with the original cause preserved.
"""

from __future__ import annotations

import hashlib
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from lxml import etree

from . import _ooxml
from ._models import (
    AnyEditOperation,
    ContentItem,
    DeletePara,
    DeleteText,
    EditResult,
    InsertParaAfter,
    InsertParaBefore,
    InsertTextAfter,
    InsertTextBefore,
    OperationResult,
    ReplacePara,
    ReplaceText,
    Selector,
)
from ._schema import validate_batch
from .anchors import AnchorManifest
from .content import normalize
from .errors import (
    BatchOperationError,
    DocxKnifeError,
    InvalidPatternError,
    ValidationError,
)
from .paragraph import (
    apply_delete_text,
    apply_insert_text_after,
    apply_insert_text_before,
    apply_replace_text,
)
from .textmap import compile_selector

if TYPE_CHECKING:
    from .document import Document

_PREVIEW_LIMIT = 80
_CHANGE_LOG_LIMIT = 4096


def _preview(text: str) -> str:
    if len(text) <= _PREVIEW_LIMIT:
        return text
    return text[: _PREVIEW_LIMIT - 1] + "\u2026"


# ---------------------------------------------------------------------------
# Normalized operations and execution plan
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class _Normalized:
    """One input operation plus its position in the original array."""

    sequence_index: int
    op: AnyEditOperation


@dataclass(slots=True)
class _AnchorGroup:
    """Every operation targeting one anchor id, in input order."""

    anchor_id: str
    first_seq: int
    inserts_before: list[_Normalized] = field(default_factory=list)
    inserts_after: list[_Normalized] = field(default_factory=list)
    replace: _Normalized | None = None
    delete: _Normalized | None = None
    text_ops: list[_Normalized] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class _Snapshot:
    root_xml: bytes
    ids_in_order: tuple[str, ...]
    next_counter: int
    paragraph_hashes: Mapping[str, bytes]


# ---------------------------------------------------------------------------
# Executor
# ---------------------------------------------------------------------------


class BatchExecutor:
    """One-shot atomic batch runner."""

    def __init__(
        self,
        document: Document,
        operations: Sequence[AnyEditOperation] | Iterable[AnyEditOperation],
        *,
        normalize_text: bool = False,
        envelope: Mapping[str, Any] | None = None,
    ) -> None:
        self._doc = document
        self._ops: tuple[AnyEditOperation, ...] = tuple(operations)
        self._normalize_text = normalize_text
        self._envelope = envelope

    def run(self) -> EditResult:
        # 1. Prevalidate everything that does not need I/O.
        self._prevalidate()

        # 2. Conflict detection + execution plan.
        groups, order = self._build_plan()

        # 3. Snapshot.
        snapshot = self._snapshot()
        pre_ids = set(snapshot.ids_in_order)

        # 4. Execute.
        results_by_seq: dict[int, OperationResult] = {}
        pending_log: list[dict[str, Any]] = []
        try:
            expected_delta = self._execute_plan(
                groups, order, results_by_seq, pending_log
            )
            self._precommit_validate(snapshot, pre_ids, groups, expected_delta)
        except BaseException as exc:
            self._rollback(snapshot)
            if isinstance(exc, BatchOperationError):
                self._audit_failure(exc, snapshot)
                raise
            if isinstance(exc, DocxKnifeError):
                target_index, target_op_id = self._failure_context(results_by_seq)
                wrapped = BatchOperationError(
                    operation_index=target_index,
                    op_id=target_op_id,
                    reason=str(exc),
                    cause=exc,
                    rolled_back=True,
                )
                self._audit_failure(wrapped, snapshot)
                raise wrapped from exc
            raise

        # 5. Commit: flush pending log entries.
        self._flush_log(pending_log)

        ordered = tuple(results_by_seq[i] for i in range(len(self._ops)))
        return EditResult(results=ordered)

    # -------------------------------------------------------- prevalidation

    def _prevalidate(self) -> None:
        if self._envelope is not None:
            validate_batch(self._envelope)

        if not self._ops:
            raise ValidationError(
                stage="prevalidation",
                checks=("nonempty",),
                failed_check="nonempty",
            )

        manifest_ids = set(self._doc._manifest.ordered_ids())

        for idx, op in enumerate(self._ops):
            if isinstance(op, (InsertParaBefore, InsertParaAfter, ReplacePara)):
                if not op.items:
                    raise BatchOperationError(
                        operation_index=idx,
                        op_id=op.op_id,
                        reason=f"{op.op}: items must be non-empty",
                    )
                for item_idx, item in enumerate(op.items):
                    if (item.content_literal is None) == (item.content_ref is None):
                        raise BatchOperationError(
                            operation_index=idx,
                            op_id=op.op_id,
                            reason=(
                                f"{op.op}: item[{item_idx}] must provide exactly "
                                "one of content_literal or content_ref"
                            ),
                        )
                if op.target_id not in manifest_ids:
                    raise BatchOperationError(
                        operation_index=idx,
                        op_id=op.op_id,
                        reason=f"{op.op}: unknown target_id {op.target_id!r}",
                    )
            elif isinstance(op, DeletePara):
                if not op.target_ids:
                    raise BatchOperationError(
                        operation_index=idx,
                        op_id=op.op_id,
                        reason="delete_para: target_ids must be non-empty",
                    )
                if len(set(op.target_ids)) != len(op.target_ids):
                    raise BatchOperationError(
                        operation_index=idx,
                        op_id=op.op_id,
                        reason="delete_para: target_ids must be unique",
                    )
                for tid in op.target_ids:
                    if tid not in manifest_ids:
                        raise BatchOperationError(
                            operation_index=idx,
                            op_id=op.op_id,
                            reason=f"delete_para: unknown target_id {tid!r}",
                        )
            elif isinstance(op, (ReplaceText, InsertTextBefore, InsertTextAfter)):
                if (op.content_literal is None) == (op.content_ref is None):
                    raise BatchOperationError(
                        operation_index=idx,
                        op_id=op.op_id,
                        reason=(
                            f"{op.op}: exactly one of content_literal or "
                            "content_ref is required"
                        ),
                    )
                self._validate_text_op(idx, op, manifest_ids)
            elif isinstance(op, DeleteText):
                self._validate_text_op(idx, op, manifest_ids)
            else:  # pragma: no cover - defensive
                raise BatchOperationError(
                    operation_index=idx,
                    op_id=getattr(op, "op_id", ""),
                    reason=f"unsupported operation type: {type(op).__name__}",
                )

    def _validate_text_op(
        self,
        idx: int,
        op: ReplaceText | DeleteText | InsertTextBefore | InsertTextAfter,
        manifest_ids: set[str],
    ) -> None:
        if op.target_id not in manifest_ids:
            raise BatchOperationError(
                operation_index=idx,
                op_id=op.op_id,
                reason=f"{op.op}: unknown target_id {op.target_id!r}",
            )
        occurrence = op.occurrence
        if occurrence is not None:
            if isinstance(occurrence, bool) or not isinstance(occurrence, int):
                raise BatchOperationError(
                    operation_index=idx,
                    op_id=op.op_id,
                    reason=f"{op.op}: occurrence must be an integer or None",
                )
            if occurrence < -1:
                raise BatchOperationError(
                    operation_index=idx,
                    op_id=op.op_id,
                    reason=f"{op.op}: occurrence must be >= -1",
                )
        try:
            compile_selector(op.find)
        except InvalidPatternError as exc:
            raise BatchOperationError(
                operation_index=idx,
                op_id=op.op_id,
                reason=f"{op.op}: invalid selector: {exc.reason}",
                cause=exc,
            ) from exc

    # ------------------------------------------------------- planning

    def _build_plan(self) -> tuple[dict[str, _AnchorGroup], list[str]]:
        groups: dict[str, _AnchorGroup] = {}

        def _group(anchor_id: str, seq: int) -> _AnchorGroup:
            g = groups.get(anchor_id)
            if g is None:
                g = _AnchorGroup(anchor_id=anchor_id, first_seq=seq)
                groups[anchor_id] = g
            return g

        for seq, op in enumerate(self._ops):
            norm = _Normalized(sequence_index=seq, op=op)
            if isinstance(op, InsertParaBefore):
                _group(op.target_id, seq).inserts_before.append(norm)
            elif isinstance(op, InsertParaAfter):
                _group(op.target_id, seq).inserts_after.append(norm)
            elif isinstance(op, ReplacePara):
                g = _group(op.target_id, seq)
                if g.replace is not None:
                    raise BatchOperationError(
                        operation_index=seq,
                        op_id=op.op_id,
                        reason=(
                            f"conflict: replace_para on {op.target_id!r} "
                            "already assigned to an earlier operation"
                        ),
                    )
                g.replace = norm
            elif isinstance(op, DeletePara):
                for tid in op.target_ids:
                    g = _group(tid, seq)
                    if g.delete is not None:
                        raise BatchOperationError(
                            operation_index=seq,
                            op_id=op.op_id,
                            reason=(
                                f"conflict: delete_para on {tid!r} "
                                "already scheduled by an earlier delete_para"
                            ),
                        )
                    g.delete = norm
            else:  # text op
                _group(op.target_id, seq).text_ops.append(norm)

        # Conflict matrix — see design §9.
        for anchor_id, g in groups.items():
            if g.delete is not None and g.replace is not None:
                blame = (
                    g.replace
                    if g.replace.sequence_index > g.delete.sequence_index
                    else g.delete
                )
                raise BatchOperationError(
                    operation_index=blame.sequence_index,
                    op_id=blame.op.op_id,
                    reason=(
                        f"conflict: replace_para and delete_para both target "
                        f"{anchor_id!r}"
                    ),
                )
            if g.delete is not None and g.inserts_after:
                blame = g.inserts_after[0]
                raise BatchOperationError(
                    operation_index=blame.sequence_index,
                    op_id=blame.op.op_id,
                    reason=(
                        f"conflict: insert_para_after on {anchor_id!r} that is "
                        "also scheduled for deletion"
                    ),
                )

        # Stable order by earliest sequence index touching the anchor.
        order = sorted(groups.keys(), key=lambda aid: groups[aid].first_seq)
        return groups, order

    # ------------------------------------------------------- snapshot

    def _snapshot(self) -> _Snapshot:
        root = self._doc._root
        xml_bytes = etree.tostring(root)
        ids_in_order = self._doc._manifest.ordered_ids()
        next_counter = self._doc._manifest._next_counter
        hashes = _paragraph_hashes(self._doc)
        return _Snapshot(
            root_xml=xml_bytes,
            ids_in_order=ids_in_order,
            next_counter=next_counter,
            paragraph_hashes=hashes,
        )

    def _rollback(self, snapshot: _Snapshot) -> None:
        # Reparse from the immutable byte snapshot: any partial DOM edits
        # (including partially-mutated shared subtrees) are undone exactly.
        parser = _ooxml.build_secure_parser()
        new_root = etree.fromstring(snapshot.root_xml, parser)
        doc = self._doc
        doc._tree = etree.ElementTree(new_root)
        doc._root = new_root
        # Rebuild the manifest and rebind IDs by document order.
        new_manifest = AnchorManifest(new_root)
        ordered_new_nodes = [
            node
            for node, _loc, _style, _para_id in _ooxml.iter_editable_paragraphs(new_root)
        ]
        if len(ordered_new_nodes) != len(snapshot.ids_in_order):  # pragma: no cover
            raise ValidationError(
                stage="rollback",
                checks=("ids_in_order_length",),
                failed_check="ids_in_order_length",
            )
        for target_id, node in zip(
            snapshot.ids_in_order, ordered_new_nodes, strict=True
        ):
            new_manifest.bind(target_id, node)
        new_manifest._next_counter = snapshot.next_counter
        doc._manifest = new_manifest
        doc._invalidate_index()
        doc._last_op_warnings = ()

    # ------------------------------------------------------- execution

    def _execute_plan(
        self,
        groups: dict[str, _AnchorGroup],
        order: list[str],
        results_by_seq: dict[int, OperationResult],
        pending_log: list[dict[str, Any]],
    ) -> int:
        doc = self._doc
        expected_delta = 0
        seen_deletes: set[int] = set()

        for anchor_id in order:
            g = groups[anchor_id]

            # 1. insert_para_before — each op individually keeps original
            #    anchor A stationary; item arrays accumulate before A in
            #    operation order.
            for norm in sorted(g.inserts_before, key=lambda n: n.sequence_index):
                op = norm.op
                assert isinstance(op, InsertParaBefore)
                before_preview = _preview(doc.get_paragraph(anchor_id))
                handles = doc.insert_para_before(
                    anchor_id,
                    list(op.items),
                    raw=op.raw,
                    normalize_text=self._normalize_text,
                )
                new_ids = [h.id for h in handles]
                previews = [_preview(doc.get_paragraph(nid)) for nid in new_ids]
                results_by_seq[norm.sequence_index] = OperationResult(
                    op_id=op.op_id,
                    op=op.op,
                    target_id=anchor_id,
                    new_ids=tuple(new_ids),
                    before_preview=before_preview,
                    after_previews=tuple(previews),
                )
                pending_log.append(
                    _insert_log_entry(op.op_id, op.op, anchor_id, before_preview, new_ids, previews)
                )
                expected_delta += len(new_ids)

            # 2. insert_para_after — the anchor for each successive op is the
            #    last inserted paragraph so item order builds up naturally
            #    (A,B,C,D,E,...) without after-insertion reversal.
            after_anchor = anchor_id
            for norm in sorted(g.inserts_after, key=lambda n: n.sequence_index):
                op = norm.op
                assert isinstance(op, InsertParaAfter)
                before_preview = _preview(doc.get_paragraph(anchor_id))
                handles = doc.insert_para_after(
                    after_anchor,
                    list(op.items),
                    raw=op.raw,
                    normalize_text=self._normalize_text,
                )
                new_ids = [h.id for h in handles]
                if new_ids:
                    after_anchor = new_ids[-1]
                previews = [_preview(doc.get_paragraph(nid)) for nid in new_ids]
                results_by_seq[norm.sequence_index] = OperationResult(
                    op_id=op.op_id,
                    op=op.op,
                    target_id=anchor_id,
                    new_ids=tuple(new_ids),
                    before_preview=before_preview,
                    after_previews=tuple(previews),
                )
                pending_log.append(
                    _insert_log_entry(op.op_id, op.op, anchor_id, before_preview, new_ids, previews)
                )
                expected_delta += len(new_ids)

            # 3. Text ops in input order.
            for norm in sorted(g.text_ops, key=lambda n: n.sequence_index):
                results_by_seq[norm.sequence_index], entry = self._exec_text_op(
                    norm, anchor_id
                )
                pending_log.append(entry)

            # 4. replace_para (after all insertions on the same anchor).
            if g.replace is not None:
                result, entry, added = self._exec_replace_para(g.replace, anchor_id)
                results_by_seq[g.replace.sequence_index] = result
                pending_log.append(entry)
                expected_delta += added - 1

            # 5. delete_para — the multi-target op fires once at its
            #    scheduled anchor visit; guard against duplicate visits when
            #    the same op targets several anchors.
            if g.delete is not None:
                seq = g.delete.sequence_index
                if seq not in seen_deletes:
                    seen_deletes.add(seq)
                    del_op = g.delete.op
                    assert isinstance(del_op, DeletePara)
                    previews = [
                        _preview(doc.get_paragraph(tid)) for tid in del_op.target_ids
                    ]
                    doc.delete_para(list(del_op.target_ids))
                    results_by_seq[seq] = OperationResult(
                        op_id=del_op.op_id,
                        op=del_op.op,
                        target_ids=tuple(del_op.target_ids),
                        before_preview=previews[0] if previews else None,
                    )
                    pending_log.append({
                        "op_id": del_op.op_id,
                        "op": del_op.op,
                        "target_ids": list(del_op.target_ids),
                        "status": "success",
                        "warnings": [],
                        "before": {"previews": previews},
                        "after": {"deleted_count": len(del_op.target_ids)},
                    })
                    expected_delta -= len(del_op.target_ids)

        return expected_delta

    def _exec_text_op(
        self, norm: _Normalized, anchor_id: str
    ) -> tuple[OperationResult, dict[str, Any]]:
        doc = self._doc
        op = norm.op
        elem = doc._manifest.resolve(anchor_id)
        selector: Selector = op.find  # type: ignore[union-attr]
        occurrence = getattr(op, "occurrence", None)
        if isinstance(op, ReplaceText):
            replacement = self._resolve_text_content(op)
            outcome = apply_replace_text(
                elem, selector, replacement, occurrence, target_id=anchor_id
            )
        elif isinstance(op, DeleteText):
            outcome = apply_delete_text(
                elem, selector, occurrence, target_id=anchor_id
            )
        elif isinstance(op, InsertTextBefore):
            text = self._resolve_text_content(op)
            outcome = apply_insert_text_before(
                elem, selector, text, occurrence, target_id=anchor_id
            )
        elif isinstance(op, InsertTextAfter):
            text = self._resolve_text_content(op)
            outcome = apply_insert_text_after(
                elem, selector, text, occurrence, target_id=anchor_id
            )
        else:  # pragma: no cover - defensive
            raise BatchOperationError(
                operation_index=norm.sequence_index,
                op_id=op.op_id,
                reason=f"unsupported text op: {type(op).__name__}",
            )
        result = OperationResult(
            op_id=op.op_id,
            op=op.op,
            target_id=anchor_id,
            warnings=outcome.warnings,
            before_preview=outcome.before_preview,
            after_previews=(outcome.after_preview,),
        )
        log_entry = {
            "op_id": op.op_id,
            "op": op.op,
            "target_id": anchor_id,
            "status": "success",
            "warnings": list(outcome.warnings),
            "before": {"preview": outcome.before_preview},
            "after": {"preview": outcome.after_preview},
        }
        return result, log_entry

    def _resolve_text_content(
        self,
        op: ReplaceText | InsertTextBefore | InsertTextAfter,
    ) -> str:
        from .content import resolve_items

        item = ContentItem(
            content_literal=op.content_literal, content_ref=op.content_ref
        )
        resolved = resolve_items((item,), raw=False, config=self._doc._content_config)
        # Text ops accept a single string; a literal containing paragraph
        # breaks joins with \n so it round-trips as line breaks inside one
        # paragraph.
        text = "\n".join(resolved[0].paragraphs)
        if self._normalize_text:
            text = normalize(text)
        return text

    def _exec_replace_para(
        self, norm: _Normalized, anchor_id: str
    ) -> tuple[OperationResult, dict[str, Any], int]:
        doc = self._doc
        op = norm.op
        assert isinstance(op, ReplacePara)
        before_preview = _preview(doc.get_paragraph(anchor_id))
        handles = doc.replace_para(
            anchor_id,
            list(op.items),
            raw=op.raw,
            normalize_text=self._normalize_text,
        )
        warnings = doc.last_operation_warnings
        new_ids = [h.id for h in handles]
        previews = [_preview(doc.get_paragraph(nid)) for nid in new_ids]
        result = OperationResult(
            op_id=op.op_id,
            op=op.op,
            target_id=anchor_id,
            new_ids=tuple(new_ids),
            warnings=warnings,
            before_preview=before_preview,
            after_previews=tuple(previews),
        )
        entry = {
            "op_id": op.op_id,
            "op": op.op,
            "target_id": anchor_id,
            "status": "success",
            "warnings": list(warnings),
            "before": {"preview": before_preview},
            "after": {
                "inserted_count": len(new_ids),
                "previews": previews,
                "new_ids": list(new_ids),
            },
        }
        return result, entry, len(new_ids)

    # -------------------------------------------------------- precommit

    def _precommit_validate(
        self,
        snapshot: _Snapshot,
        pre_ids: set[str],
        groups: dict[str, _AnchorGroup],
        expected_delta: int,
    ) -> None:
        doc = self._doc

        # 1. main XML round-trips.
        try:
            etree.tostring(doc._root)
        except (etree.XMLSyntaxError, ValueError) as exc:  # pragma: no cover
            raise ValidationError(
                stage="commit",
                checks=("xml_serializable",),
                failed_check="xml_serializable",
            ) from exc

        # 2. Paragraph count delta matches expectation.
        actual_count = doc.paragraph_count()
        expected_count = len(snapshot.ids_in_order) + expected_delta
        if actual_count != expected_count:
            raise ValidationError(
                stage="commit",
                checks=("paragraph_count_delta",),
                failed_check="paragraph_count_delta",
            )

        # 3. Deleted ids gone; replace targets gone.
        current_ids = set(doc._manifest.ordered_ids())
        for anchor_id, g in groups.items():
            if g.delete is not None and anchor_id in current_ids:
                raise ValidationError(
                    stage="commit",
                    checks=("deletion_removed",),
                    failed_check="deletion_removed",
                )
            if g.replace is not None and anchor_id in current_ids:
                raise ValidationError(
                    stage="commit",
                    checks=("replace_removed_old_id",),
                    failed_check="replace_removed_old_id",
                )

        # 4. Untouched-paragraph canonical fidelity.
        touched_ids = _touched_ids(groups)
        current_hashes = _paragraph_hashes(doc)
        for tid in pre_ids - touched_ids:
            before = snapshot.paragraph_hashes.get(tid)
            after = current_hashes.get(tid)
            if before is None or after is None or before != after:
                raise ValidationError(
                    stage="commit",
                    checks=("untouched_fidelity",),
                    failed_check="untouched_fidelity",
                )

    # -------------------------------------------------------- audit / log

    def _flush_log(self, pending: list[dict[str, Any]]) -> None:
        log = self._doc._change_log
        for entry in pending:
            if len(log) >= _CHANGE_LOG_LIMIT:
                return
            log.append(entry)

    def _audit_failure(self, error: BatchOperationError, snapshot: _Snapshot) -> None:
        log = self._doc._change_log
        if len(log) >= _CHANGE_LOG_LIMIT:
            return
        cause = error.__cause__
        log.append({
            "batch": True,
            "status": "failed",
            "error_type": type(cause).__name__ if cause is not None else type(error).__name__,
            "operation_index": error.operation_index,
            "op_id": error.op_id,
            "reason": _preview(error.reason),
            "rolled_back": True,
            "expected_paragraph_count": len(snapshot.ids_in_order),
            "actual_paragraph_count": self._doc.paragraph_count(),
        })

    def _failure_context(
        self,
        results_by_seq: dict[int, OperationResult],
    ) -> tuple[int, str]:
        # First op without a recorded result — that is where execution stopped.
        for idx in range(len(self._ops)):
            if idx not in results_by_seq:
                return idx, self._ops[idx].op_id
        last = len(self._ops) - 1
        return last, self._ops[last].op_id


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _insert_log_entry(
    op_id: str,
    op: str,
    anchor_id: str,
    before_preview: str,
    new_ids: list[str],
    previews: list[str],
) -> dict[str, Any]:
    return {
        "op_id": op_id,
        "op": op,
        "target_id": anchor_id,
        "status": "success",
        "warnings": [],
        "before": {"preview": before_preview},
        "after": {
            "inserted_count": len(new_ids),
            "previews": previews,
            "new_ids": list(new_ids),
        },
    }


def _touched_ids(groups: dict[str, _AnchorGroup]) -> set[str]:
    touched: set[str] = set()
    for anchor_id, g in groups.items():
        if (
            g.inserts_before
            or g.inserts_after
            or g.text_ops
            or g.replace is not None
            or g.delete is not None
        ):
            touched.add(anchor_id)
    return touched


def _paragraph_hashes(doc: Document) -> dict[str, bytes]:
    """SHA-256 of each paragraph's serialized XML keyed by manifest id.

    ``etree.tostring`` on a subelement includes parent-scoped namespace
    declarations, so two structurally identical paragraphs from different
    contexts hash the same regardless of the ancestor chain that wrapped them
    at snapshot time versus commit time.
    """
    hashes: dict[str, bytes] = {}
    for target_id in doc._manifest.ordered_ids():
        node = doc._manifest.resolve(target_id)
        raw = etree.tostring(node)
        hashes[target_id] = hashlib.sha256(raw).digest()
    return hashes


__all__ = ["BatchExecutor"]
