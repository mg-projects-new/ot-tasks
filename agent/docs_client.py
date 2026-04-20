"""Google Docs reader: fetch tab text from a doc referenced in a Jira brief.

Strategy:
- Extract doc ID from a URL in the brief's Description
- Fetch document metadata (all tabs)
- Pick the tab that best matches the ticket Summary (fuzzy token overlap)
- Return plain text of that tab
"""

from __future__ import annotations

import logging
import re
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

# URL patterns for Google Docs references in Jira wiki markup
DOC_URL_RE = re.compile(
    r"docs\.google\.com/document/d/([a-zA-Z0-9_-]{20,})",
    re.IGNORECASE,
)


def extract_doc_id(text: str) -> str | None:
    """Find the first Google Doc ID referenced in a chunk of text (e.g. Jira Description)."""
    if not text:
        return None
    m = DOC_URL_RE.search(text)
    return m.group(1) if m else None


def _normalize_for_match(s: str) -> str:
    """Lowercase, drop punctuation, collapse whitespace — for fuzzy title matching."""
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _score_title_match(ticket_summary: str, tab_title: str) -> float:
    """Return [0..1] similarity score between the ticket summary and a tab title.

    We strip SMM-specific prefixes from the summary (like 'SMM / 8.05 /
    First steps in trading /') and then compute a mix of:
    - token overlap ratio
    - SequenceMatcher ratio
    """
    # Strip common Jira summary prefixes to get the "real" topic
    cleaned = ticket_summary
    # Remove everything before the last '/' if present (e.g., "SMM / 8.05 / First steps / Topic" -> "Topic")
    if "/" in cleaned:
        cleaned = cleaned.rsplit("/", 1)[-1]

    a = _normalize_for_match(cleaned)
    b = _normalize_for_match(tab_title)
    if not a or not b:
        return 0.0

    # Token overlap
    a_tokens = set(a.split())
    b_tokens = set(b.split())
    if a_tokens and b_tokens:
        overlap = len(a_tokens & b_tokens) / len(a_tokens | b_tokens)
    else:
        overlap = 0.0

    # Sequence similarity
    seq = SequenceMatcher(None, a, b).ratio()

    # Weighted blend
    return 0.6 * overlap + 0.4 * seq


def _read_content_elements(content: list) -> list[str]:
    """Recursively extract text from a 'content' array of a Docs body or tab."""
    out: list[str] = []
    for element in content or []:
        if "paragraph" in element:
            para = element["paragraph"]
            pieces: list[str] = []
            for el in para.get("elements", []):
                tr = el.get("textRun")
                if tr and "content" in tr:
                    pieces.append(tr["content"])
            if pieces:
                out.append("".join(pieces))
        elif "table" in element:
            for row in element["table"].get("tableRows", []):
                for cell in row.get("tableCells", []):
                    out.extend(_read_content_elements(cell.get("content", [])))
        elif "sectionBreak" in element:
            # ignore section breaks
            pass
    return out


def _flatten_tabs(tabs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Tabs can be nested (child tabs). Return a flat list of {title, body_content}."""
    out: list[dict[str, Any]] = []
    for tab in tabs or []:
        props = tab.get("tabProperties", {}) or {}
        title = props.get("title") or "(untitled)"
        doc_tab = tab.get("documentTab") or {}
        body = doc_tab.get("body") or {}
        content = body.get("content") or []
        out.append({"title": title, "content": content})
        # Recurse into child tabs
        out.extend(_flatten_tabs(tab.get("childTabs") or []))
    return out


def fetch_tab_text_for_ticket(
    agent_dir: Path, *, doc_id: str, ticket_summary: str,
    min_score: float = 0.25,
) -> tuple[str, str] | None:
    """Fetch the Google Doc, find the tab best matching ticket_summary, return (tab_title, plain_text).

    Returns None if no tab crosses the min_score threshold — caller should fall back
    to generating without expert source.
    """
    from sheets_client import _get_credentials  # reuse existing OAuth flow
    creds = _get_credentials(agent_dir)
    docs = build("docs", "v1", credentials=creds, cache_discovery=False)

    try:
        # includeTabsContent=True returns all tabs with their bodies in one call
        doc = docs.documents().get(
            documentId=doc_id,
            includeTabsContent=True,
        ).execute()
    except Exception as e:
        logger.warning("Could not fetch doc %s: %s", doc_id, e)
        return None

    tabs = doc.get("tabs") or []
    flat = _flatten_tabs(tabs)

    if not flat:
        # Legacy docs without tabs: whole body is the content
        body = doc.get("body") or {}
        content = body.get("content") or []
        if content:
            text = "".join(_read_content_elements(content)).strip()
            if text:
                return (doc.get("title", "(whole doc)"), text)
        return None

    # Score each tab against the ticket summary
    scored: list[tuple[float, str, list]] = []
    for t in flat:
        score = _score_title_match(ticket_summary, t["title"])
        scored.append((score, t["title"], t["content"]))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_title, best_content = scored[0]

    logger.info(
        "Doc tab match: best='%s' (score=%.2f). Runner-ups: %s",
        best_title, best_score,
        [(t[1], round(t[0], 2)) for t in scored[1:4]],
    )

    if best_score < min_score:
        logger.warning(
            "No tab matched ticket summary '%s' well enough (best=%s at %.2f). "
            "Returning None so caller can skip expert-source injection.",
            ticket_summary, best_title, best_score,
        )
        return None

    text = "".join(_read_content_elements(best_content)).strip()
    if not text:
        logger.warning("Best-matched tab '%s' has no text content.", best_title)
        return None

    # Cap length to keep prompt size reasonable
    max_chars = 15000
    if len(text) > max_chars:
        text = text[:max_chars] + "\n\n[... expert content truncated ...]"

    return (best_title, text)
