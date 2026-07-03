"""Shared DOCX builders for benchmark modules."""

from __future__ import annotations

import zipfile
from pathlib import Path

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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


def _wrap(body_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{_W_NS}">'
        f"<w:body>{body_xml}</w:body>"
        "</w:document>"
    )


def _write(path: Path, body_xml: str) -> Path:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/document.xml", _wrap(body_xml))
    return path


def build_many_paragraphs(path: Path, n: int, *, keyword_every: int = 7) -> Path:
    """Write a DOCX with ``n`` single-run paragraphs, sprinkling a keyword."""
    parts: list[str] = []
    for i in range(n):
        text = f"Paragraph {i} content line."
        if i % keyword_every == 0:
            text += " KEYWORD"
        parts.append(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")
    return _write(path, "".join(parts))


def build_multi_run_paragraph(path: Path, run_count: int) -> Path:
    """Write a DOCX containing one paragraph split into ``run_count`` runs."""
    runs = "".join(f"<w:r><w:t>seg{i:03d}</w:t></w:r>" for i in range(run_count))
    body = f"<w:p>{runs}</w:p>"
    return _write(path, body)


__all__ = ["build_many_paragraphs", "build_multi_run_paragraph"]
