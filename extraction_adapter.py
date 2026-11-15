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
    if value is not None:
        return value, float(confidence) if confidence is not None else None
    if expected_type is str:
        return "", float(confidence) if confidence is not None else None
    if expected_type is int:
        return 0, 0.0
    if expected_type is float:
        return 0.0, 0.0
    raise TypeError(f"Unsupported non-nullable schema type: {expected_type!r}")


def confidence_map(records: Iterable[Mapping[str, Any]]) -> Dict[str, float | None]:
    """Return confidence as a non-breaking sidecar keyed by occurrence index.

    Occurrence indexes preserve duplicate labels instead of collapsing them.
    """
    result: Dict[str, float | None] = {}
    counts: Dict[str, int] = {}
    for record in records:
        label = str(record.get("label") or "")
        index = counts.get(label, 0)
        counts[label] = index + 1
        key = f"{label}#{index}"
        score = record.get("confidence_score")
        result[key] = float(score) if score is not None else None
    return result
