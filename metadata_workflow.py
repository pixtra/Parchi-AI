"""Exhaustive metadata, label matching, and template helpers."""
from __future__ import annotations

import json
from typing import Any

def sheet_text(workbook: Any) -> list[dict[str, Any]]:
    """Return one text unit per worksheet, including hidden worksheets."""
    pages: list[dict[str, Any]] = []
    for number, sheet in enumerate(workbook.worksheets, start=1):
        lines: list[str] = []
        for row in sheet.iter_rows(values_only=True):
            values = ["" if value is None else str(value) for value in row]
            if any(value.strip() for value in values):
                lines.append(" | ".join(values))
        text = "\n".join(lines)
        pages.append({
            "page_number": number,
            "sheet_name": sheet.title,
            "text": text,
            "readable": bool(text.strip()),
        })
    return pages


def json_bytes(payload: Any) -> bytes:
    return json.dumps(payload, indent=2, ensure_ascii=False, allow_nan=False).encode("utf-8")