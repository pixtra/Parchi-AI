"""Matching rules for applying an existing template.

Per extracted label, in this order:
1. A saved decision for this template (exact match, approved suggestion, or a choice the user
   made, including "not mapped") is always respected and never overridden.
2. An exact label match (case-insensitive, surrounding whitespace ignored) is mapped automatically.
3. Only for labels with no exact match, RapidFuzz-style similarity suggestions (>= 70%) are
   offered, and each one needs explicit approval or rejection.
4. Anything else is reported as unmatched. Nothing is forced.
"""
from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Sequence, Set, Tuple

from metadata_workflow import exact_template_label, suggestions

SUGGESTION_MIN_SCORE = 70


def plan_template_matches(
    source_labels: Sequence[str],
    template_labels: Sequence[str],
    decisions: Mapping[str, Optional[str]],
    rejected: Set[Tuple[str, str]],
) -> Dict[str, List[Any]]:
    """Plan the matching for one template.

    Returns:
    - "decided": (source, target-or-None) pairs that already have a saved decision;
    - "auto": (source, target) exact matches to save now;
    - "review": suggestion dicts (source, target, score) that need a decision;
    - "unmatched": source labels with no exact match and no suggestion.
    """
    template_labels = [str(label) for label in template_labels]
    plan: Dict[str, List[Any]] = {"decided": [], "auto": [], "review": [], "unmatched": []}
    for source in source_labels:
        if source in decisions:
            plan["decided"].append((source, decisions[source]))
            continue
        exact = exact_template_label(source, template_labels)
        if exact is not None:
            plan["auto"].append((source, exact))
            continue
        matches = [
            match for match in suggestions(source, template_labels, limit=3, min_score=SUGGESTION_MIN_SCORE)
            if (source, match["label"]) not in rejected
        ]
        if matches:
            for match in matches:
                plan["review"].append({"source": source, "target": match["label"], "score": match["score"]})
        else:
            plan["unmatched"].append(source)
    return plan