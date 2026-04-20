"""Anthropic Messages API wrapper with retries and JSON extraction."""
from __future__ import annotations
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Any
import anthropic
from anthropic import APIConnectionError, APIStatusError
logger = logging.getLogger(__name__)
DEFAULT_MODEL = "claude-sonnet-4-5"
MAX_TOKENS = 16000


def normalize_api_key(k: str) -> str:
    return (k or "").strip().strip('"').strip("'")


def load_env_key(agent_dir: Path) -> str:
    from dotenv import load_dotenv
    load_dotenv(agent_dir / ".env", override=True)
    return normalize_api_key(os.environ.get("ANTHROPIC_API_KEY", ""))


def get_model_name() -> str:
    env_model = (os.environ.get("OT_AGENT_MODEL") or "").strip()
    return env_model or DEFAULT_MODEL


def _extract_balanced_json(s: str) -> str | None:
    """Find the first balanced {...} block, tracking string state so that braces
    inside string literals don't mess up the depth count. Returns the substring
    or None if no balanced block found."""
    start = s.find("{")
    if start == -1:
        return None
    depth = 0
    in_string = False
    escape_next = False
    for i in range(start, len(s)):
        ch = s[i]
        if escape_next:
            escape_next = False
            continue
        if ch == "\\" and in_string:
            escape_next = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return s[start : i + 1]
    return None


def _repair_json_strings(s: str) -> str:
    """Best-effort repair of common JSON emission issues:
    - unescaped raw newlines/tabs inside string literals -> \\n / \\t
    - Windows line endings normalized
    - smart quotes inside keys/structural positions left alone
    Operates only when *inside a JSON string literal*, tracked by state.
    """
    out: list[str] = []
    in_string = False
    escape_next = False
    for ch in s:
        if escape_next:
            out.append(ch)
            escape_next = False
            continue
        if ch == "\\":
            out.append(ch)
            if in_string:
                escape_next = True
            continue
        if ch == '"':
            in_string = not in_string
            out.append(ch)
            continue
        if in_string:
            if ch == "\n":
                out.append("\\n")
            elif ch == "\r":
                out.append("\\r")
            elif ch == "\t":
                out.append("\\t")
            else:
                out.append(ch)
        else:
            out.append(ch)
    return "".join(out)


def _strip_trailing_commas(s: str) -> str:
    """Remove trailing commas before ] or } (outside of strings)."""
    result: list[str] = []
    in_string = False
    escape_next = False
    i = 0
    while i < len(s):
        ch = s[i]
        if escape_next:
            result.append(ch)
            escape_next = False
            i += 1
            continue
        if ch == "\\" and in_string:
            result.append(ch)
            escape_next = True
            i += 1
            continue
        if ch == '"':
            in_string = not in_string
            result.append(ch)
            i += 1
            continue
        if not in_string and ch == ",":
            # peek ahead: skip whitespace, if next non-ws is ] or }, drop the comma
            j = i + 1
            while j < len(s) and s[j] in " \t\n\r":
                j += 1
            if j < len(s) and s[j] in "]}":
                i += 1
                continue
        result.append(ch)
        i += 1
    return "".join(result)


def extract_json_object(text: str) -> dict[str, Any]:
    """
    Parse JSON from model output. Strips optional ```json fences and leading junk.
    Tries multiple recovery strategies if naive parsing fails.
    """
    s = text.strip()
    fence = re.match(r"^```(?:json)?\s*\n?(.*?)\n?```\s*$", s, re.DOTALL | re.IGNORECASE)
    if fence:
        s = fence.group(1).strip()

    # Strategy 1: parse as-is
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass

    # Strategy 2: extract balanced {...} block using string-aware parser
    block = _extract_balanced_json(s)
    if block is None:
        raise ValueError("No JSON object found in model response")

    try:
        return json.loads(block)
    except json.JSONDecodeError:
        pass

    # Strategy 3: repair common issues (unescaped newlines inside strings)
    repaired = _repair_json_strings(block)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass

    # Strategy 4: also strip trailing commas
    cleaned = _strip_trailing_commas(repaired)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        # Show a snippet around the error for debugging
        err_pos = getattr(e, "pos", 0)
        start = max(0, err_pos - 80)
        end = min(len(cleaned), err_pos + 80)
        snippet = cleaned[start:end]
        raise ValueError(
            f"Could not parse model response as JSON after repairs. "
            f"Error: {e.msg} at pos {err_pos}. Context: ...{snippet!r}..."
        ) from e


def call_claude_json(
    *,
    agent_dir: Path,
    system_prompt: str,
    user_prompt: str,
) -> dict[str, Any]:
    """
    Send system + user messages; return parsed JSON dict.
    On 401: prints one line and re-raises a clean error (no traceback from caller).
    Retries once on 5xx / connection errors with short backoff. No retry on 401.
    Retries once with a cleanup instruction if JSON parsing fails.
    """
    key = load_env_key(agent_dir)
    if not key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is missing. Copy agent/.env.example to agent/.env and set your key."
        )
    env_model = (os.environ.get("OT_AGENT_MODEL") or "").strip()
    model_name = env_model or DEFAULT_MODEL
    client = anthropic.Anthropic(api_key=key)

    # Strengthen system prompt with an explicit JSON discipline reminder
    strict_system = system_prompt + (
        "\n\nIMPORTANT FOR JSON OUTPUT: Your entire response must be a single valid JSON object. "
        "Inside string values, escape every newline as \\n, every tab as \\t, every double quote as \\\". "
        "Do NOT include raw unescaped newlines inside string values. "
        "No markdown fences, no prose before or after the JSON."
    )

    def _once(system_text: str, user_text: str) -> str:
        msg = client.messages.create(
            model=model_name,
            max_tokens=MAX_TOKENS,
            system=system_text,
            messages=[{"role": "user", "content": user_text}],
        )
        parts: list[str] = []
        for block in msg.content:
            if hasattr(block, "text"):
                parts.append(block.text)
        return "".join(parts)

    last_err: Exception | None = None
    last_text: str | None = None
    for attempt in range(2):
        try:
            text = _once(strict_system, user_prompt)
            last_text = text
            try:
                return extract_json_object(text)
            except ValueError as parse_err:
                # JSON parse failure: retry once with a repair request
                if attempt == 0:
                    logger.warning("First JSON parse failed: %s — retrying with repair prompt", parse_err)
                    repair_user = (
                        f"Your previous response could not be parsed as JSON because of a formatting error. "
                        f"Please re-emit the SAME content as a single strictly valid JSON object. "
                        f"Escape every newline as \\n, every tab as \\t, every double quote as \\\". "
                        f"No markdown fences. No prose.\n\n"
                        f"Original task was:\n{user_prompt}"
                    )
                    try:
                        retry_text = _once(strict_system, repair_user)
                        last_text = retry_text
                        return extract_json_object(retry_text)
                    except Exception as retry_err:
                        last_err = retry_err
                        continue
                raise
        except APIStatusError as e:
            if getattr(e, "status_code", None) == 401:
                print(
                    "Anthropic rejected the key (401). Check console.anthropic.com → API Keys. "
                    "Ensure billing is funded."
                )
                raise RuntimeError("Anthropic 401 Unauthorized") from e
            last_err = e
            if attempt == 0 and getattr(e, "status_code", 0) >= 500:
                time.sleep(2.0)
                continue
            raise
        except APIConnectionError as e:
            last_err = e
            if attempt == 0:
                time.sleep(2.0)
                continue
            raise
    if last_err:
        raise last_err
    raise RuntimeError("call_claude_json failed without capturing an error")