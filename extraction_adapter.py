"""Backwards-compatible adaptation of nullable model extraction values."""
from __future__ import annotations

from typing import Any, Dict, Iterable, Mapping, Tuple, Type


def coerce_nullable_value(
    value: Any,
    expected_type: Type[Any],
    confidence: Any = None,
) -> Tuple[Any, float | None]:
    """Coerce a nullable model value without changing the original schema.

    Numeric nulls also invalidate the confidence for that field.  The optional
    confidence return is a sidecar; callers that only need the original value
    can ignore it.
    """
    # A real value already satisfies the original schema, so preserve it.
    if value is not None:
        return value, float(confidence) if confidence is not None else None
    # Text fields use an empty string because the legacy schema is non-nullable.
    if expected_type is str:
        return "", float(confidence) if confidence is not None else None
    # A missing number is represented by zero and receives zero confidence.
    if expected_type is int:
        return 0, 0.0
    if expected_type is float:
        return 0.0, 0.0
    # Fail loudly for unsupported schema types instead of silently corrupting data.
    raise TypeError(f"Unsupported non-nullable schema type: {expected_type!r}")


def confidence_map(records: Iterable[Mapping[str, Any]]) -> Dict[str, float | None]:
    """Return confidence as a non-breaking sidecar keyed by occurrence index.

    Occurrence indexes preserve duplicate labels instead of collapsing them.
    """
    # The result is keyed by label plus occurrence number to preserve duplicates.
    result: Dict[str, float | None] = {}
    counts: Dict[str, int] = {}
    for record in records:
        # Normalize only the key; the original record remains unchanged.
        label = str(record.get("label") or "")
        index = counts.get(label, 0)
        counts[label] = index + 1
        key = f"{label}#{index}"
        # Keep None when Gemini did not provide a confidence score.
        score = record.get("confidence_score")
        result[key] = float(score) if score is not None else None
    return result
