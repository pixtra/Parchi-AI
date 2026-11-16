"""Exhaustive metadata, label matching, and template helpers."""
from __future__ import annotations

import json
from typing import Any, Iterable

try:
    from rapidfuzz import fuzz, process
except ImportError:  # pragma: no cover - optional acceleration
    fuzz = None
    process = None


def sheet_text(workbook: Any) -> list[dict[str, Any]]:
    """Return one text unit per worksheet, including hidden worksheets."""
    # Keep one output object per worksheet so sheet identity survives extraction.
    pages: list[dict[str, Any]] = []
    for number, sheet in enumerate(workbook.worksheets, start=1):
        # Convert each non-empty spreadsheet row into a readable pipe-separated line.
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


def unique_labels(records: Iterable[dict[str, Any]]) -> list[str]:
    """Return non-empty labels once, sorted without case sensitivity."""
    # A set removes exact duplicates before the stable human-friendly sort.
    return sorted({str(record.get("label") or "").strip()
                   for record in records if str(record.get("label") or "").strip()},
                  key=str.casefold)


def normalize_label(label: Any) -> str:
    """Normalize a label for case-insensitive matching."""
    return str(label or "").strip().casefold()


def exact_template_label(source_label: str, requested_labels: Iterable[str]) -> str | None:
    """Return the requested label whose normalized spelling exactly matches the source."""
    # Return the requested spelling so exports use the client's label text.
    normalized = normalize_label(source_label)
    return next(
        (str(label) for label in requested_labels if normalize_label(label) == normalized),
        None,
    )


def suggestions(
    source_label: str,
    requested_labels: Iterable[str],
    limit: int = 5,
    min_score: float = 70,
) -> list[dict[str, Any]]:
    """Return the closest requested labels for a source label.

    RapidFuzz is used when installed because it is faster.  The standard-library
    fallback keeps the application usable in a minimal environment.
    """
    # Remove blank candidates before asking either matcher to score them.
    labels = [str(label) for label in requested_labels if str(label).strip()]
    if not labels:
        return []
    # Keep the user's original spelling in the returned suggestions.
    query = str(source_label).strip()
    if process is not None:
        # RapidFuzz returns (candidate, score, position) tuples.
        return [
            {"label": match[0], "score": round(float(match[1]), 1)}
            for match in process.extract(
                query, labels, scorer=fuzz.ratio, limit=limit, score_cutoff=min_score
            )
        ]
    # Import lazily because this branch is only needed without RapidFuzz.
    from difflib import SequenceMatcher
    # Score every candidate against the normalized source label.
    return sorted(
        ({"label": label, "score": round(100 * SequenceMatcher(
            None, query.casefold(), label.casefold()).ratio(), 1)}
         for label in labels
         if 100 * SequenceMatcher(None, query.casefold(), label.casefold()).ratio() >= min_score),
        key=lambda item: item["score"], reverse=True,
    )[:limit]


def json_bytes(payload: Any) -> bytes:
    """Serialize export data as readable UTF-8 JSON bytes."""
    # Reject NaN and infinity so exported JSON is valid for every consumer.
    return json.dumps(payload, indent=2, ensure_ascii=False, allow_nan=False).encode("utf-8")
