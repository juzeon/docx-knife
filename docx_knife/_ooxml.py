"""OOXML namespace constants and lxml helpers.

Populated in Phase 2. This module is intentionally minimal today; it establishes
the import surface for later phases without leaking implementation details.
"""

from __future__ import annotations

from typing import Final

W_NS: Final[str] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS: Final[str] = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_NS: Final[str] = "http://www.w3.org/XML/1998/namespace"

NSMAP: Final[dict[str, str]] = {
    "w": W_NS,
    "w14": W14_NS,
}


def qn(tag: str) -> str:
    """Expand a namespace-prefixed tag (``w:p`` -> ``{...}p``)."""
    prefix, _, local = tag.partition(":")
    if not local:
        return tag
    ns = NSMAP.get(prefix)
    if ns is None:
        raise KeyError(f"unknown xml namespace prefix: {prefix!r}")
    return f"{{{ns}}}{local}"
