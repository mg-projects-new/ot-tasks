#!/usr/bin/env python3
"""OT Copywriting Agent — generate English SMM copy from Jira (or xlsx), save locally and/or to Google Sheets."""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook

AGENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = AGENT_DIR.parent
sys.path.insert(0, str(AGENT_DIR))

from claude_client import call_claude_json, get_model_name  # noqa: E402
from excel_io.fill_template import (  # noqa: E402
    serialize_workbook_for_prompt,
    slugify_summary,
)
from excel_io.jira_export import read_jira_brief  # noqa: E402
from jira_client import fetch_backlog_tickets, fetch_ticket_full, ticket_to_brief  # noqa: E402
from rubric_detector import detect_rubric  # noqa: E402
from rubrics import glossary as glossary_rubric  # noqa: E402
from rubrics import strategy as strategy_rubric  # noqa: E402
from rubrics import ugc as ugc_rubric  # noqa: E402

IMPLEMENTED = {"Strategy", "Glossary", "UGC"}
SCAFFOLDED = {"First steps in trading"}

logger = logging.getLogger("ot_agent")

TEXT_SECTION_RE = re.compile(
    r"h3\.\s*\*?\s*TEXT\*?:?\s*(.*?)(?=\n\s*h3\.\s|\n\s*----|\Z)",
    re.DOTALL | re.IGNORECASE,
)


def extract_text_section(description: str) -> str:
    if not description:
        return ""
    m = TEXT_SECTION_RE.search(description)
    return m.group(1).strip() if m else ""


def load_config() -> dict[str, Any]:
    path = AGENT_DIR / "config.yaml"
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def load_base_system_prompt() -> str:
    return (AGENT_DIR / "prompts" / "base_prompt.md").read_text(encoding="utf-8")


def sanitize_brief(brief: dict[str, Any]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    cap = 120_000
    for k, v in brief.items():
        if v is None:
            continue
        s = str(v)
        if len(s) > cap:
            s = s[:cap] + "\n...[truncated]"
        out[k] = s
    return out


def jira_browse_url(issue_key: str) -> str:
    return f"https://space307.atlassian.net/browse/{issue_key}"


def collect_task_files(tasks_dir: Path, single: Path | None) -> list[Path]:
    if single is not None:
        p = single if single.is_absolute() else PROJECT_ROOT / single
        if not p.is_file():
            raise FileNotFoundError(p)
        return [p]
    return sorted(
        x for x in tasks_dir.glob("*.xlsx")
        if x.is_file() and not x.name.startswith("~$")
    )


def serialize_examples_dir(examples_dir: Path) -> str:
    parts: list[str] = []
    for p in sorted(examples_dir.glob("*.xlsx")):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        parts.append(serialize_workbook_for_prompt(p))
    return "\n\n".join(parts)


def append_run_log(project_root: Path, lines: list[str]) -> None:
    log_path = project_root / "Completed" / "_run_log.md"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    chunk = [f"## Run {ts}", ""] + lines + [""]
    with open(log_path, "a", encoding="utf-8") as f:
        f.write("\n".join(chunk))


def build_user_prompt(
    *, rubric: str, rubric_notes: str, brief: dict[str, Any],
    schema_text: str, examples_text: str, extra_context: str = "",
) -> str:
    text_section = extract_text_section(str(brief.get("Description", "")))
    mandatory_block = ""
    if text_section:
        mandatory_block = (
            "\n=== MANDATORY INSTRUCTIONS (from Jira `h3. TEXT` section) ===\n"
            "These instructions OVERRIDE the reference examples. If they specify\n"
            "character limits, CTAs, tone, or format rules, follow them literally.\n"
            "---\n"
            f"{text_section}\n"
            "---\n"
        )
    extra_block = f"\n{extra_context}\n" if extra_context else ""
    brief_json = json.dumps(sanitize_brief(brief), ensure_ascii=False, indent=2)
    return f"""Rubric: {rubric}
{mandatory_block}{extra_block}
Rubric-specific notes:
{rubric_notes}

Jira brief (JSON — field names are export headers; Description uses Jira wiki markup):
{brief_json}

Output JSON schema (every value is an English string; no locale dicts):
{schema_text}

Reference examples (serialized Excel workbooks — for VOICE and LAYOUT reference only, NOT content):
{examples_text}

Strict rules:
- Reply with one JSON object only. No markdown fences, no commentary.
- English copy only. No translations, no non-English text.
- Follow the MANDATORY INSTRUCTIONS above if present; they override example patterns.
"""


def process_strategy(*, brief, examples_dir, rubric_notes, dry_run):
    schema_text = strategy_rubric.strategy_schema_json_text()
    ex_text = serialize_examples_dir(examples_dir) or "(No example workbooks found.)"
    user = build_user_prompt(rubric="Strategy", rubric_notes=rubric_notes, brief=brief, schema_text=schema_text, examples_text=ex_text)
    system = load_base_system_prompt()
    if dry_run:
        return strategy_rubric.strategy_json_template()
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return strategy_rubric.normalize_strategy_payload(raw)


def process_glossary(*, brief, examples_dir, rubric_notes, dry_run):
    layout = glossary_rubric.get_glossary_layout(examples_dir)
    schema_text = glossary_rubric.glossary_schema_json_text(layout)
    ex_text = serialize_examples_dir(examples_dir) or "(No example workbooks found.)"
    user = build_user_prompt(rubric="Glossary", rubric_notes=rubric_notes, brief=brief, schema_text=schema_text, examples_text=ex_text)
    system = load_base_system_prompt()
    if dry_run:
        return glossary_rubric.glossary_json_template(), layout
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return glossary_rubric.normalize_glossary_payload(raw, layout), layout


def process_ugc(*, brief, examples_dir, rubric_notes, dry_run, override_count):
    plan = ugc_rubric.plan_from_brief(brief, override=override_count)
    if not plan.videos:
        raise RuntimeError(
            "UGC: could not detect any videos in the brief. "
            "Either the TEXT section is missing `*Видео N.*` markers, or pass `--videos N`."
        )
    logger.info("UGC: planning %d video(s)", len(plan.videos))
    schema_text = ugc_rubric.ugc_schema_json_text(plan)
    ex_text = serialize_examples_dir(examples_dir) or "(No example workbooks found.)"
    lines = ["=== VIDEOS TO PRODUCE ==="]
    for v in plan.videos:
        sid = f" (Script {v.script_id})" if v.script_id else ""
        lines.append(f"- Video {v.video_number}{sid}")
        if v.raw_block:
            lines.append(f"  Brief context:\n{v.raw_block[:2000]}\n")
    lines.append("=== END VIDEOS ===")
    extra = "\n".join(lines)
    user = build_user_prompt(rubric="UGC", rubric_notes=rubric_notes, brief=brief, schema_text=schema_text, examples_text=ex_text, extra_context=extra)
    system = load_base_system_prompt()
    if dry_run:
        return ugc_rubric.ugc_json_template(plan), plan
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return ugc_rubric.normalize_ugc_payload(raw, plan), plan


def save_strategy_output(data, *, out_path, jira_url):
    wb = strategy_rubric.build_strategy_workbook(data, jira_url=jira_url)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def save_glossary_output(data, layout, *, examples_dir, out_path, jira_url):
    wb = glossary_rubric.create_minimal_glossary_workbook()
    glossary_rubric.fill_glossary_workbook(wb, data, jira_url=jira_url, layout=layout)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    wb.close()


def save_ugc_output(data, *, out_path, jira_url):
    wb = ugc_rubric.build_ugc_workbook(data, jira_url=jira_url)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------- Google Sheets delivery ----------

def upload_to_sheets(rubric: str, data: Any, *, title: str, jira_url: str) -> str | None:
    """Create a native Google Sheet for this rubric, return its URL or None on failure."""
    from sheets_client import (  # noqa: WPS433 — lazy import so default mode doesn't require google libs
        build_glossary_sheet, build_strategy_sheet, build_ugc_sheet, get_output_folder_id,
    )
    from rubrics.glossary import ROW_ORDER as GLOSSARY_ROWS, ROW_DISPLAY as GLOSSARY_DISPLAY
    from rubrics.strategy import ROW_ORDER as STRATEGY_ROWS

    try:
        folder_id = get_output_folder_id()
    except RuntimeError as e:
        logger.error("Sheets upload skipped: %s", e)
        return None

    try:
        if rubric == "Glossary":
            _, url = build_glossary_sheet(
                AGENT_DIR, title=title, jira_url=jira_url, data=data,
                row_order=GLOSSARY_ROWS, row_display=GLOSSARY_DISPLAY, folder_id=folder_id,
            )
        elif rubric == "Strategy":
            _, url = build_strategy_sheet(
                AGENT_DIR, title=title, jira_url=jira_url, data=data,
                row_order=STRATEGY_ROWS, folder_id=folder_id,
            )
        elif rubric == "UGC":
            videos = data.get("videos") or []
            _, url = build_ugc_sheet(
                AGENT_DIR, title=title, jira_url=jira_url, videos=videos, folder_id=folder_id,
            )
        else:
            return None
        return url
    except Exception as e:
        logger.exception("Google Sheets upload failed for %s", rubric)
        print(f"     ⚠ Sheets upload failed: {e}")
        return None


# ---------- Per-ticket dispatcher ----------

def process_one_brief(
    *, brief: dict[str, Any], rubric: str, dry_run: bool, cfg: dict[str, Any],
    override_videos: int | None = None, to_sheets: bool = False,
) -> tuple[bool, str, str | None]:
    """Returns (ok, message, sheet_url_or_none)."""
    rubrics_cfg = cfg.get("rubrics") or {}
    notes = (rubrics_cfg.get(rubric) or {}).get("notes") or ""
    examples_dir = PROJECT_ROOT / f"{rubric} Examples"
    completed_dir = PROJECT_ROOT / "Completed" / rubric

    issue_key = str(brief.get("Issue key") or "").strip()
    summary = brief.get("Summary") or ""
    if not issue_key:
        return False, "no Issue key in brief", None

    slug = slugify_summary(str(summary) if summary else "task")
    out_name = f"{issue_key}_{slug}.xlsx"
    out_path = completed_dir / out_name

    if any(completed_dir.glob(f"{issue_key}_*.xlsx")):
        return False, f"skipped (already in Completed/{rubric}/)", None

    jira_url = jira_browse_url(issue_key)
    sheet_url: str | None = None

    if rubric == "Strategy":
        data = process_strategy(brief=brief, examples_dir=examples_dir, rubric_notes=notes, dry_run=dry_run)
        if not dry_run:
            save_strategy_output(data, out_path=out_path, jira_url=jira_url)
            if to_sheets:
                sheet_url = upload_to_sheets("Strategy", data, title=f"{issue_key} — {summary}", jira_url=jira_url)
    elif rubric == "Glossary":
        data, layout = process_glossary(brief=brief, examples_dir=examples_dir, rubric_notes=notes, dry_run=dry_run)
        if not dry_run:
            save_glossary_output(data, layout, examples_dir=examples_dir, out_path=out_path, jira_url=jira_url)
            if to_sheets:
                sheet_url = upload_to_sheets("Glossary", data, title=f"{issue_key} — {summary}", jira_url=jira_url)
    elif rubric == "UGC":
        data, _ = process_ugc(brief=brief, examples_dir=examples_dir, rubric_notes=notes, dry_run=dry_run, override_count=override_videos)
        if not dry_run:
            save_ugc_output(data, out_path=out_path, jira_url=jira_url)
            if to_sheets:
                sheet_url = upload_to_sheets("UGC", data, title=f"{issue_key} — {summary}", jira_url=jira_url)
    else:
        return False, f"rubric '{rubric}' has no handler", None

    msg = f"dry-run (would save {out_name})" if dry_run else f"✓ saved {out_name}"
    return True, msg, sheet_url


# ---------- Modes ----------

def run_from_jira(args, cfg: dict[str, Any]) -> int:
    logger.info("Fetching Jira backlog (assignee=currentUser, status=Backlog)...")
    try:
        tickets = fetch_backlog_tickets(AGENT_DIR)
    except RuntimeError as e:
        print(str(e)); return 2
    except Exception as e:
        logger.exception("Jira fetch failed")
        print(f"Jira fetch failed: {e}"); return 2

    if not tickets:
        print("No tickets in your backlog. Nothing to do."); return 0

    print(f"Fetched {len(tickets)} ticket(s) from Jira backlog.")
    log_lines: list[str] = []
    logger.info("Model: %s", get_model_name())

    for t in tickets:
        key = t.get("key") or ""
        summary = (t.get("fields") or {}).get("summary") or ""
        label = f"{key} — {summary}"

        rubric = detect_rubric(summary)
        if rubric is None:
            print(f"  ○ {label}\n     skipped: couldn't detect rubric from summary")
            log_lines.append(f"- [unknown] {key} — skipped (rubric not detected: {summary!r})")
            continue
        if rubric in SCAFFOLDED:
            print(f"  ○ {label}\n     skipped: rubric '{rubric}' is scaffolded only — not implemented yet")
            log_lines.append(f"- [{rubric}] {key} — skipped (rubric not implemented)")
            continue
        if rubric not in IMPLEMENTED:
            print(f"  ○ {label}\n     skipped: rubric '{rubric}' is not known")
            log_lines.append(f"- [{rubric}] {key} — skipped (rubric unknown)")
            continue

        try:
            full = fetch_ticket_full(AGENT_DIR, key)
            brief = ticket_to_brief(full)
        except Exception as e:
            logger.exception("Failed fetching %s", key)
            print(f"  ✗ {label}\n     error fetching ticket: {e}")
            log_lines.append(f"- [{rubric}] {key} — error fetching ticket ({e})")
            continue

        print(f"  • {label}\n     rubric: {rubric}")
        try:
            ok, msg, sheet_url = process_one_brief(
                brief=brief, rubric=rubric, dry_run=args.dry_run, cfg=cfg,
                override_videos=args.videos, to_sheets=args.to_sheets,
            )
            status = "✓" if ok else "○"
            print(f"     {status} {msg}")
            if sheet_url:
                print(f"     🔗 {sheet_url}")
            log_lines.append(f"- [{rubric}] {key} — {msg}" + (f" — {sheet_url}" if sheet_url else ""))
        except RuntimeError as e:
            if "401" in str(e):
                if log_lines: append_run_log(PROJECT_ROOT, log_lines)
                return 1
            logger.exception("Failed %s", key)
            log_lines.append(f"- [{rubric}] {key} — error ({e})")
        except Exception as e:
            logger.exception("Failed %s", key)
            print(f"     ✗ error: {e}")
            log_lines.append(f"- [{rubric}] {key} — error ({e})")

    if log_lines:
        append_run_log(PROJECT_ROOT, log_lines)
    return 0


def run_from_files(args, cfg: dict[str, Any]) -> int:
    rubric = (args.rubric or "").strip()
    if rubric not in IMPLEMENTED:
        if rubric in SCAFFOLDED:
            print(f"Rubric '{rubric}' is scaffolded only."); return 2
        print(f"Rubric '{rubric}' is not implemented. Implemented: {', '.join(sorted(IMPLEMENTED))}."); return 2

    tasks_dir = PROJECT_ROOT / f"{rubric} Tasks"
    if not tasks_dir.is_dir():
        print(f"Tasks folder not found: {tasks_dir}"); return 2

    single: Path | None = Path(args.input) if args.input else None
    try:
        files = collect_task_files(tasks_dir, single)
    except FileNotFoundError as e:
        print(e); return 2
    if not files:
        print(f"No .xlsx task files in {tasks_dir}"); return 1

    log_lines: list[str] = []
    logger.info("Model: %s", get_model_name())

    for path in files:
        try:
            rel = path.relative_to(PROJECT_ROOT)
        except ValueError:
            rel = path
        try:
            brief = read_jira_brief(path)
        except Exception as e:
            logger.exception("Failed to read %s", path)
            log_lines.append(f"- [{rubric}] {rel.name} — error reading brief ({e})")
            continue
        try:
            ok, msg, sheet_url = process_one_brief(
                brief=brief, rubric=rubric, dry_run=args.dry_run, cfg=cfg,
                override_videos=args.videos, to_sheets=args.to_sheets,
            )
            prefix = "[dry-run] " if args.dry_run else ""
            print(f"{prefix}{rel.name}: {msg}")
            if sheet_url:
                print(f"  🔗 {sheet_url}")
            log_lines.append(f"- [{rubric}] {rel.name} — {msg}" + (f" — {sheet_url}" if sheet_url else ""))
        except RuntimeError as e:
            if "401" in str(e):
                if log_lines: append_run_log(PROJECT_ROOT, log_lines)
                return 1
            logger.exception("Failed %s", rel.name)
            log_lines.append(f"- [{rubric}] {rel.name} — error ({e})")
        except Exception as e:
            logger.exception("Failed %s", rel.name)
            log_lines.append(f"- [{rubric}] {rel.name} — error ({e})")

    if log_lines:
        append_run_log(PROJECT_ROOT, log_lines)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="OT Copywriting Agent")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--from-jira", action="store_true",
                      help="Fetch tickets from Jira (backlog, assigned to you) and auto-route by rubric")
    mode.add_argument("--rubric", help='Rubric name for file mode (e.g., "Glossary")')

    parser.add_argument("--all", action="store_true")
    parser.add_argument("--input", type=str)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--videos", type=int, default=None, help="UGC only: override number of videos")
    parser.add_argument("--to-sheets", action="store_true",
                        help="Also create a native Google Sheet in your Drive folder (in addition to local xlsx)")

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )

    cfg = load_config()
    if args.from_jira:
        return run_from_jira(args, cfg)
    if not args.all and not args.input:
        print("File mode requires --all or --input"); return 2
    return run_from_files(args, cfg)


if __name__ == "__main__":
    sys.exit(main())