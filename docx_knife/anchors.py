"""Anchor manifest: paragraph-ID allocation and node bookkeeping (Phase 2).

The manifest is the only surface where paragraph IDs cross the LLM boundary.
It stores live ``lxml`` nodes and enforces:

* IDs are allocated monotonically as ``p_000001``, ``p_000002``, ...
* Invalidated IDs are never reused (the allocator counter only advances).
* ``resolve`` fails closed when the bound node has been detached from the
  document root; the manifest never relocates by text, index, or similarity.

Snapshots serialize the mapping so that Phase 6 batch transactions can roll
back the manifest alongside the DOM.
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from .errors import ParagraphNotFoundError

_ID_TEMPLATE = "p_{:06d}"


@dataclass(frozen=True, slots=True)
class AnchorSnapshot:
    """Immutable snapshot of an ``AnchorManifest`` for transactional rollback."""

    order: tuple[str, ...]
    bindings: tuple[tuple[str, etree._Element], ...]
    next_counter: int


class AnchorManifest:
    """Maps stable paragraph IDs to live ``<w:p>`` nodes."""

    def __init__(self, root: etree._Element) -> None:
        self._root = root
        self._by_id: dict[str, etree._Element] = {}
        self._by_node: dict[int, str] = {}
        self._order: list[str] = []
        self._next_counter = 1

    @property
    def root(self) -> etree._Element:
        return self._root

    def allocate(self) -> str:
        """Return the next monotonic ID; the counter never rewinds."""
        target_id = _ID_TEMPLATE.format(self._next_counter)
        self._next_counter += 1
        return target_id

    def bind(self, target_id: str, node: etree._Element) -> None:
        """Bind ``target_id`` to ``node``. Overwrites any prior binding for that ID."""
        if target_id in self._by_id:
            self._by_node.pop(id(self._by_id[target_id]), None)
        else:
            self._order.append(target_id)
        self._by_id[target_id] = node
        self._by_node[id(node)] = target_id

    def id_for_node(self, node: etree._Element) -> str | None:
        return self._by_node.get(id(node))

    def invalidate(self, target_id: str) -> None:
        node = self._by_id.pop(target_id, None)
        if node is not None:
            self._by_node.pop(id(node), None)
            self._order.remove(target_id)

    def resolve(self, target_id: str) -> etree._Element:
        node = self._by_id.get(target_id)
        if node is None:
            raise ParagraphNotFoundError(target_id=target_id)
        if not self._reachable(node):
            raise ParagraphNotFoundError(target_id=target_id)
        return node

    def ordered_ids(self) -> tuple[str, ...]:
        return tuple(self._order)

    def __contains__(self, target_id: object) -> bool:
        return isinstance(target_id, str) and target_id in self._by_id

    def __len__(self) -> int:
        return len(self._by_id)

    def snapshot(self) -> AnchorSnapshot:
        return AnchorSnapshot(
            order=tuple(self._order),
            bindings=tuple((tid, self._by_id[tid]) for tid in self._order),
            next_counter=self._next_counter,
        )

    def restore(self, snapshot: AnchorSnapshot) -> None:
        self._order = list(snapshot.order)
        self._by_id = dict(snapshot.bindings)
        self._by_node = {id(node): tid for tid, node in snapshot.bindings}
        self._next_counter = snapshot.next_counter

    def _reachable(self, node: etree._Element) -> bool:
        cur: etree._Element | None = node
        while cur is not None:
            if cur is self._root:
                return True
            cur = cur.getparent()
        return False


__all__ = ["AnchorManifest", "AnchorSnapshot"]
