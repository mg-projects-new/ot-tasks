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


def normalize_api_key(raw: str | None) -> str:
    """Trim whitespace, strip BOM, remove surrounding quotes — do not alter key body."""
    if raw is None:
        return ""
    s = raw.strip()
    if s.startswith("\ufeff"):
        s = s.lstrip("\ufeff")
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1]
    return s.strip()


def load_env_key(agent_dir: Path) -> str:
    """Load ANTHROPIC_API_KEY from agent/.env via python-dotenv."""
    from dotenv import load_dotenv

    env_path = agent_dir / ".env"
    load_dotenv(env_path)
    return normalize_api_key(os.environ.get("ANTHROPIC_API_KEY", ""))


def extract_json_object(text: str) -> dict[str, Any]:
    """
    Parse JSON from model output. Strips optional ```json fences and leading junk.
    """
    s = text.strip()
    fence = re.match(r"^```(?:json)?\s*\n?(.*?)\n?```\s*$", s, re.DOTALL | re.IGNORECASE)
    if fence:
        s = fence.group(1).strip()
    # First { ... } balanced block as fallback
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass
    start = s.find("{")
    if start == -1:
        raise ValueError("No JSON object found in model response")
    depth = 0
    for i, ch in enumerate(s[start:], start=start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(s[start : i + 1])
    raise ValueError("Could not parse balanced JSON object from model response")


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
    """
    key = load_env_key(agent_dir)
    if not key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is missing. Copy agent/.env.example to agent/.env and set your key."
        )

    env_model = (os.environ.get("OT_AGENT_MODEL") or "").strip()
    model_name = env_model or DEFAULT_MODEL

    client = anthropic.Anthropic(api_key=key)

    def _once() -> str:
        msg = client.messages.create(
            model=model_name,
            max_tokens=MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        parts: list[str] = []
        for block in msg.content:
            if hasattr(block, "text"):
                parts.append(block.text)
        return "".join(parts)

    last_err: Exception | None = None
    for attempt in range(2):
        try:
            text = _once()
            return extract_json_object(text)
        except APIStatusError as e:
            if getattr(e, "status_code", None) == 401:
                print(
                    "Anthropic rejected the key (401). Check console.anthropic.com → API Keys. "
                    "Ensure billing is funded."
                )
                raise RuntimeError("Anthropic API 401") from None
            if getattr(e, "status_code", None) and 500 <= e.status_code < 600 and attempt == 0:
                wait = 2.0**attempt
                logger.warning("Anthropic %s, retrying in %.1fs", e.status_code, wait)
                time.sleep(wait)
                last_err = e
                continue
            raise
        except (OSError, APIConnectionError) as e:
            if attempt == 0:
                logger.warning("Connection error, retrying once: %s", e)
                time.sleep(2.0)
                last_err = e
                continue
            raise
    if last_err:
        raise last_err
    raise RuntimeError("Unreachable")


def get_model_name() -> str:
    env_model = (os.environ.get("OT_AGENT_MODEL") or "").strip()
    return env_model or DEFAULT_MODEL
