"""Detect which rubric a Jira ticket belongs to, based on its Summary field."""

from __future__ import annotations

# Detection rules in order of priority.
# Each tuple: (rubric_name, list of keywords — any match triggers the rubric)
RULES: list[tuple[str, list[str]]] = [
    ("UGC", ["ugc", "tiktok"]),
    ("Glossary", ["glossary"]),
    ("First steps in trading", ["first steps", "first step"]),
    ("Strategy", ["strategy", "indicator", "pattern"]),
]


def detect_rubric(summary: str) -> str | None:
    """Return the rubric name based on keywords in the summary, or None if unknown."""
    if not summary:
        return None
    s = summary.lower()
    for rubric, keywords in RULES:
        for kw in keywords:
            if kw in s:
                return rubric
    return None