#!/usr/bin/env python3
"""OT Copywriting Agent — generate English SMM copy from Jira task exports."""

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
from rubrics import glossary as glossary_rubric  # noqa: E402
from rubrics import strategy as strategy_rubric  # noqa: E402
from rubrics import ugc as ugc_rubric  # noqa: E402

IMPLEMENTED = {"Strategy", "Glossary", "UGC"}

logger = logging.getLogger("ot_agent")


# Matches an "h3. TEXT" section up to the next h3./---- divider or EOF
TEXT_SECTION_RE = re.compile(
    r"h3\.\s*\*?\s*TEXT\*?:?\s*(.*?)(?=\n\s*h3\.\s|\n\s*----|\Z)",
    re.DOTALL | re.IGNORECASE,
)


def extract_text_section(description: str) -> str:
    """Pull the h3. TEXT block from a Jira Description (mandatory instructions)."""
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
    *,
    rubric: str,
    rubric_notes: str,
    brief: dict[str, Any],
    schema_text: str,
    examples_text: str,
    extra_context: str = "",
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


def process_strategy(
    *, brief: dict[str, Any], examples_dir: Path, rubric_notes: str, dry_run: bool,
) -> dict[str, Any]:
    schema_text = strategy_rubric.strategy_schema_json_text()
    ex_text = serialize_examples_dir(examples_dir)
    if not ex_text.strip():
        logger.warning("No .xlsx examples in %s — model has no format reference.", examples_dir)
        ex_text = "(No example workbooks found.)"
    user = build_user_prompt(
        rubric="Strategy",
        rubric_notes=rubric_notes,
        brief=brief,
        schema_text=schema_text,
        examples_text=ex_text,
    )
    system = load_base_system_prompt()
    if dry_run:
        logger.info("Dry-run: Strategy prompt %s chars.", len(user))
        return strategy_rubric.strategy_json_template()
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return strategy_rubric.normalize_strategy_payload(raw)


def process_glossary(
    *, brief: dict[str, Any], examples_dir: Path, rubric_notes: str, dry_run: bool,
) -> tuple[dict[str, Any], glossary_rubric.GlossaryLayout]:
    layout = glossary_rubric.get_glossary_layout(examples_dir)
    schema_text = glossary_rubric.glossary_schema_json_text(layout)
    ex_text = serialize_examples_dir(examples_dir)
    if not ex_text.strip():
        logger.warning("No .xlsx examples in %s — using minimal template.", examples_dir)
        ex_text = "(No example workbooks found.)"
    user = build_user_prompt(
        rubric="Glossary",
        rubric_notes=rubric_notes,
        brief=brief,
        schema_text=schema_text,
        examples_text=ex_text,
    )
    system = load_base_system_prompt()
    if dry_run:
        logger.info("Dry-run: Glossary prompt %s chars.", len(user))
        return glossary_rubric.glossary_json_template(), layout
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return glossary_rubric.normalize_glossary_payload(raw, layout), layout


def process_ugc(
    *, brief: dict[str, Any], examples_dir: Path, rubric_notes: str, dry_run: bool,
    override_count: int | None,
) -> tuple[dict[str, Any], ugc_rubric.UGCBriefPlan]:
    plan = ugc_rubric.plan_from_brief(brief, override=override_count)
    if not plan.videos:
        raise RuntimeError(
            "UGC: could not detect any videos in the brief. "
            "Either the TEXT section is missing `*Видео N.*` markers, or pass `--videos N`."
        )
    logger.info("UGC: planning %d video(s)", len(plan.videos))

    schema_text = ugc_rubric.ugc_schema_json_text(plan)
    ex_text = serialize_examples_dir(examples_dir)
    if not ex_text.strip():
        logger.warning("No .xlsx examples in %s — model has no format reference.", examples_dir)
        ex_text = "(No example workbooks found.)"

    # Extra context block that tells Claude the exact video list
    lines = ["=== VIDEOS TO PRODUCE ==="]
    for v in plan.videos:
        sid = f" (Script {v.script_id})" if v.script_id else ""
        lines.append(f"- Video {v.video_number}{sid}")
        if v.raw_block:
            # Include the raw block so the LLM sees the caption direction ("Подводка к посту") and examples
            lines.append(f"  Brief context:\n{v.raw_block[:2000]}\n")
    lines.append("=== END VIDEOS ===")
    extra = "\n".join(lines)

    user = build_user_prompt(
        rubric="UGC",
        rubric_notes=rubric_notes,
        brief=brief,
        schema_text=schema_text,
        examples_text=ex_text,
        extra_context=extra,
    )
    system = load_base_system_prompt()
    if dry_run:
        logger.info("Dry-run: UGC prompt %s chars.", len(user))
        return ugc_rubric.ugc_json_template(plan), plan
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return ugc_rubric.normalize_ugc_payload(raw, plan), plan


def save_strategy_output(data: dict[str, Any], *, out_path: Path, jira_url: str) -> None:
    wb = strategy_rubric.build_strategy_workbook(data, jira_url=jira_url)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def save_glossary_output(
    data: dict[str, Any],
    layout: glossary_rubric.GlossaryLayout,
    *, examples_dir: Path, out_path: Path, jira_url: str,
) -> None:
    wb = glossary_rubric.create_minimal_glossary_workbook()
    glossary_rubric.fill_glossary_workbook(wb, data, jira_url=jira_url, layout=layout)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    wb.close()


def save_ugc_output(data: dict[str, Any], *, out_path: Path, jira_url: str) -> None:
    wb = ugc_rubric.build_ugc_workbook(data, jira_url=jira_url)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="OT Copywriting Agent")
    parser.add_argument("--rubric", required=True)
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--all", action="store_true")
    g.add_argument("--input", type=str)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--videos", type=int, default=None,
                        help="UGC only: override number of videos to generate")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )

    rubric = args.rubric.strip()
    if rubric not in IMPLEMENTED:
        if rubric == "First steps in trading":
            print("First steps in trading is scaffolded only: add schema in rubrics/first_steps.py after reviewing Examples.")
            return 2
        print(f"Rubric '{rubric}' is not implemented. Implemented: {', '.join(sorted(IMPLEMENTED))}.")
        return 2

    cfg = load_config()
    rubrics_cfg = cfg.get("rubrics") or {}
    notes = (rubrics_cfg.get(rubric) or {}).get("notes") or ""

    tasks_dir = PROJECT_ROOT / f"{rubric} Tasks"
    examples_dir = PROJECT_ROOT / f"{rubric} Examples"
    completed_dir = PROJECT_ROOT / "Completed" / rubric

    if not tasks_dir.is_dir():
        print(f"Tasks folder not found: {tasks_dir}")
        return 2
    if not examples_dir.is_dir():
        logger.warning("Examples folder missing: %s (continuing).", examples_dir)

    single: Path | None = Path(args.input) if args.input else None
    try:
        files = collect_task_files(tasks_dir, single)
    except FileNotFoundError as e:
        print(e)
        return 2

    if not files:
        print(f"No .xlsx task files in {tasks_dir}")
        return 1

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

        issue_key = str(brief.get("Issue key") or brief.get("Issue Key") or "").strip()
        summary = brief.get("Summary")
        if not issue_key:
            print(f"Skipping {rel.name}: no 'Issue key' in export.")
            log_lines.append(f"- [{rubric}] {rel.name} — skipped (no Issue key)")
            continue

        slug = slugify_summary(str(summary) if summary else "task")
        out_name = f"{issue_key}_{slug}.xlsx"
        out_path = completed_dir / out_name

        if any(completed_dir.glob(f"{issue_key}_*.xlsx")):
            print(f"skipping — already done: {issue_key} ({rel.name})")
            log_lines.append(f"- [{rubric}] {rel.name} — skipped (already in Completed/)")
            continue

        jira_url = jira_browse_url(issue_key)

        try:
            if rubric == "Strategy":
                data = process_strategy(
                    brief=brief, examples_dir=examples_dir,
                    rubric_notes=notes, dry_run=args.dry_run,
                )
                if not args.dry_run:
                    save_strategy_output(data, out_path=out_path, jira_url=jira_url)
            elif rubric == "Glossary":
                data, layout = process_glossary(
                    brief=brief, examples_dir=examples_dir,
                    rubric_notes=notes, dry_run=args.dry_run,
                )
                if not args.dry_run:
                    save_glossary_output(
                        data, layout,
                        examples_dir=examples_dir, out_path=out_path, jira_url=jira_url,
                    )
            elif rubric == "UGC":
                data, plan = process_ugc(
                    brief=brief, examples_dir=examples_dir,
                    rubric_notes=notes, dry_run=args.dry_run,
                    override_count=args.videos,
                )
                if not args.dry_run:
                    save_ugc_output(data, out_path=out_path, jira_url=jira_url)
        except RuntimeError as e:
            if "401" in str(e):
                return 1
            logger.exception("Failed %s", rel.name)
            log_lines.append(f"- [{rubric}] {rel.name} — error ({e})")
            continue
        except Exception as e:
            logger.exception("Failed %s", rel.name)
            log_lines.append(f"- [{rubric}] {rel.name} — error ({e})")
            continue

        try:
            rel_out = out_path.relative_to(PROJECT_ROOT)
        except ValueError:
            rel_out = out_path
        if args.dry_run:
            print(f"[dry-run] would write {rel_out}")
            log_lines.append(f"- [{rubric}] {rel.name} — dry-run (would save {out_name})")
        else:
            print(f"saved {rel_out}")
            log_lines.append(f"- [{rubric}] {rel.name} — ✓ saved {out_name}")

    if log_lines:
        append_run_log(PROJECT_ROOT, log_lines)

    return 0


if __name__ == "__main__":
    sys.exit(main())