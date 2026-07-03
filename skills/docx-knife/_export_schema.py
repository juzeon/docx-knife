"""Dump ``docx_knife._schema.BATCH_SCHEMA`` to ``agent_schema.json``.

Run once whenever the schema changes:

    python skills/docx-knife/_export_schema.py
"""

from __future__ import annotations

import json
from pathlib import Path

from docx_knife._schema import BATCH_SCHEMA


def main() -> None:
    target = Path(__file__).with_name("agent_schema.json")
    target.write_text(
        json.dumps(BATCH_SCHEMA, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
