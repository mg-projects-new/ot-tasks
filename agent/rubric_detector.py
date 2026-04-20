"""Detect rubric from a Jira ticket Summary string.

Checks keywords in priority order (more specific first).
Returns the rubric name string, or None if no rubric could be detected.
"""

from __future__ import annotations

import re


# Order matters: more specific first. Each entry: (rubric_name, [keywords])
RUBRIC_RULES: list[tuple[str, list[str]]] = [
    # Ongoing content — most specific first (Part 1 vs Part 2)
    ("Ongoing P1", [r"part\s*1", r"часть\s*1"]),
    ("Ongoing P2", [r"part\s*2", r"часть\s*2"]),
    # Other rubrics
    ("UGC", [r"\bUGC\b", r"TikTok"]),
    ("Glossary", [r"\bGlossary\b", r"глосса"]),
    ("First steps in trading", [r"first steps", r"первые шаги"]),
    ("Strategy", [r"\bStrategy\b", r"стратег", r"indicator", r"pattern"]),
]

# Gatekeepers: Ongoing P1/P2 only apply if the summary actually mentions "Онгоинг" or "Ongoing"
ONGOING_GATE = re.compile(r"(онгоинг|ongoing)", re.IGNORECASE)


def detect_rubric(summary: str) -> str | None:
    if not summary:
        return None
    s = summary.lower()
    is_ongoing = bool(ONGOING_GATE.search(summary))
    for rubric, keywords in RUBRIC_RULES:
        # Ongoing P1/P2 only match if Ongoing/Онгоинг is also in the summary
        if rubric.startswith("Ongoing") and not is_ongoing:
            continue
        for kw in keywords:
            if re.search(kw, s, re.IGNORECASE):
                return rubric
    return None
