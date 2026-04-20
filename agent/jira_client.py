"""Jira REST API client — fetch backlog tickets and convert to agent brief format."""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Any

import requests
from requests.auth import HTTPBasicAuth

logger = logging.getLogger(__name__)

JQL_BACKLOG = 'assignee = currentUser() AND status = "Backlog"'
MAX_RESULTS_PER_PAGE = 50
FIELDS_TO_FETCH = [
    "summary",
    "issuetype",
    "status",
    "assignee",
    "reporter",
    "created",
    "updated",
    "duedate",
    "priority",
    "description",
    "customfield_*",  # wildcard not supported, we ask for '*all' separately
]


def _load_jira_creds(agent_dir: Path) -> tuple[str, str, str]:
    """Read JIRA_EMAIL, JIRA_API_TOKEN, JIRA_BASE_URL from agent/.env."""
    from dotenv import load_dotenv

    load_dotenv(agent_dir / ".env")
    email = (os.environ.get("JIRA_EMAIL") or "").strip()
    token = (os.environ.get("JIRA_API_TOKEN") or "").strip()
    base = (os.environ.get("JIRA_BASE_URL") or "https://space307.atlassian.net").strip().rstrip("/")

    if not email or not token:
        raise RuntimeError(
            "Jira credentials missing. Add JIRA_EMAIL and JIRA_API_TOKEN to agent/.env. "
            "Generate a token at https://id.atlassian.com/manage-profile/security/api-tokens"
        )
    return email, token, base


def _auth(email: str, token: str) -> HTTPBasicAuth:
    return HTTPBasicAuth(email, token)


def fetch_backlog_tickets(agent_dir: Path) -> list[dict[str, Any]]:
    """Run the backlog JQL and return a list of ticket JSONs using the new /search/jql endpoint."""
    email, token, base = _load_jira_creds(agent_dir)
    url = f"{base}/rest/api/3/search/jql"
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    auth = _auth(email, token)

    all_issues: list[dict[str, Any]] = []
    next_page_token: str | None = None
    while True:
        body: dict[str, Any] = {
            "jql": JQL_BACKLOG,
            "maxResults": MAX_RESULTS_PER_PAGE,
            "fields": ["summary", "status", "assignee"],
        }
        if next_page_token:
            body["nextPageToken"] = next_page_token

        r = requests.post(url, headers=headers, json=body, auth=auth, timeout=30)
        if r.status_code == 401:
            raise RuntimeError(
                "Jira rejected credentials (401). Check JIRA_EMAIL and JIRA_API_TOKEN in agent/.env."
            )
        if r.status_code == 403:
            raise RuntimeError(
                "Jira denied access (403). Your account may not have permission for this query."
            )
        r.raise_for_status()
        data = r.json()
        issues = data.get("issues", [])
        all_issues.extend(issues)

        next_page_token = data.get("nextPageToken")
        is_last = data.get("isLast", True)
        if not next_page_token or is_last or not issues:
            break
    return all_issues


def fetch_ticket_full(agent_dir: Path, issue_key: str) -> dict[str, Any]:
    """Get full JSON for a single ticket, including description and all custom fields."""
    email, token, base = _load_jira_creds(agent_dir)
    url = f"{base}/rest/api/3/issue/{issue_key}"
    headers = {"Accept": "application/json"}
    auth = _auth(email, token)
    # expand=renderedFields gives us the rendered wiki-markup HTML + plain text
    # We want the raw ADF (atlassian document format) for the description to convert back to wiki markup
    params = {"fields": "*all", "expand": "names"}
    r = requests.get(url, headers=headers, params=params, auth=auth, timeout=30)
    r.raise_for_status()
    return r.json()


# ---------- Convert Jira API response → agent brief format ----------

def _adf_to_text(node: Any) -> str:
    """Walk an ADF (Atlassian Document Format) tree and emit a text representation
    that approximates the Jira wiki-markup your xlsx exports contained.
    This is intentionally simple; it keeps headings as 'h3. ...' and preserves line breaks."""
    if node is None:
        return ""
    if isinstance(node, str):
        return node

    out: list[str] = []
    ntype = node.get("type") if isinstance(node, dict) else None

    if ntype == "heading":
        level = node.get("attrs", {}).get("level", 3)
        inner = "".join(_adf_to_text(c) for c in node.get("content", []))
        out.append(f"\nh{level}. {inner}\n")
    elif ntype == "paragraph":
        inner = "".join(_adf_to_text(c) for c in node.get("content", []))
        out.append(inner + "\n")
    elif ntype == "text":
        text = node.get("text", "")
        marks = node.get("marks") or []
        mark_types = {m.get("type") for m in marks if isinstance(m, dict)}
        if "strong" in mark_types:
            text = f"*{text}*"
        if "em" in mark_types:
            text = f"_{text}_"
        if "link" in mark_types:
            href = ""
            for m in marks:
                if m.get("type") == "link":
                    href = (m.get("attrs") or {}).get("href", "")
            if href:
                text = f"[{text}|{href}]"
        out.append(text)
    elif ntype == "hardBreak":
        out.append("\n")
    elif ntype == "rule":
        out.append("\n----\n")
    elif ntype == "bulletList":
        for item in node.get("content", []):
            item_text = "".join(_adf_to_text(c) for c in item.get("content", []))
            out.append(f"* {item_text}")
    elif ntype == "orderedList":
        for idx, item in enumerate(node.get("content", []), start=1):
            item_text = "".join(_adf_to_text(c) for c in item.get("content", []))
            out.append(f"{idx}. {item_text}")
    elif ntype == "doc":
        for c in node.get("content", []):
            out.append(_adf_to_text(c))
    else:
        # Fallback: recurse into content
        for c in (node.get("content") or []) if isinstance(node, dict) else []:
            out.append(_adf_to_text(c))
    return "".join(out)


def ticket_to_brief(ticket_json: dict[str, Any]) -> dict[str, Any]:
    """Convert a Jira v3 API issue response into the dict shape the agent expects
    (same shape as read_jira_brief() produces from xlsx exports)."""
    fields = ticket_json.get("fields", {}) or {}
    key = ticket_json.get("key") or ""
    summary = fields.get("summary") or ""

    description_field = fields.get("description")
    if isinstance(description_field, dict):
        description = _adf_to_text(description_field).strip()
    elif isinstance(description_field, str):
        description = description_field
    else:
        description = ""

    brief: dict[str, Any] = {
        "Summary": summary,
        "Issue key": key,
        "Issue Type": (fields.get("issuetype") or {}).get("name", ""),
        "Status": (fields.get("status") or {}).get("name", ""),
        "Assignee": (fields.get("assignee") or {}).get("displayName", ""),
        "Reporter": (fields.get("reporter") or {}).get("displayName", ""),
        "Created": fields.get("created") or "",
        "Updated": fields.get("updated") or "",
        "Due date": fields.get("duedate") or "",
        "Priority": (fields.get("priority") or {}).get("name", ""),
        "Description": description,
    }

    # Custom fields — keep anything named "Custom field (Short description)" etc.
    # We use the 'names' expansion to get human-friendly names when available.
    names_map = ticket_json.get("names") or {}
    for field_id, value in fields.items():
        if not field_id.startswith("customfield_"):
            continue
        if value is None:
            continue
        label = names_map.get(field_id, field_id)
        text_value: str = ""
        if isinstance(value, dict):
            # e.g. {"value": "Common"} or a user object
            text_value = str(value.get("value") or value.get("displayName") or "")
        elif isinstance(value, list):
            if value and isinstance(value[0], dict):
                text_value = ", ".join(str(v.get("value") or v.get("displayName") or "") for v in value)
            else:
                text_value = ", ".join(str(v) for v in value)
        else:
            text_value = str(value)
        if text_value.strip():
            brief[f"Custom field ({label})"] = text_value

    return brief