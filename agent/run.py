#!/usr/bin/env python3
"""OT Copywriting Agent — generate SMM copy from Jira task exports."""

from __future__ import annotations

import argparse
import json
import logging
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
IMPLEMENTED = {"Strategy", "Glossary"}

logger = logging.getLogger("ot_agent")


def load_config() -> dict[str, Any]:
    path = AGENT_DIR / "config.yaml"
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_base_system_prompt() -> str:
    p = AGENT_DIR / "prompts" / "base_prompt.md"
    return p.read_text(encoding="utf-8")


def sanitize_brief(brief: dict[str, Any]) -> dict[str, Any]:
    """Truncate extremely long fields for the API prompt."""
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
    paths = sorted(
        x
        for x in tasks_dir.glob("*.xlsx")
        if x.is_file() and not x.name.startswith("~$")
    )
    return paths


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
) -> str:
    brief_json = json.dumps(sanitize_brief(brief), ensure_ascii=False, indent=2)
    return f"""Rubric: {rubric}

Rubric-specific notes:
{rubric_notes}

Jira brief (JSON — field names are export headers; Description uses Jira wiki markup as stored):
{brief_json}

Output JSON schema (produce valid JSON matching this structure; every leaf is a string):
{schema_text}

Reference examples (serialized Excel workbooks for tone and format):
{examples_text}

Strict rules:
- Reply with one JSON object only. No markdown fences, no commentary.
- Fill all required locales and keys; use empty strings where a cell should stay blank.
"""


def process_strategy(
    *,
    brief: dict[str, Any],
    examples_dir: Path,
    rubric_notes: str,
    dry_run: bool,
) -> dict[str, Any]:
    schema_text = strategy_rubric.strategy_schema_json_text()
    ex_text = serialize_examples_dir(examples_dir)
    if not ex_text.strip():
        logger.warning("No .xlsx examples found in %s — model has no format reference.", examples_dir)
        ex_text = "(No example workbooks found in Examples folder.)"
    user = build_user_prompt(
        rubric="Strategy",
        rubric_notes=rubric_notes,
        brief=brief,
        schema_text=schema_text,
        examples_text=ex_text,
    )
    system = load_base_system_prompt()
    if dry_run:
        logger.info("Dry-run: would call Claude with Strategy schema (%s chars user prompt).", len(user))
        return strategy_rubric.strategy_json_template()
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    return strategy_rubric.normalize_strategy_payload(raw)


def process_glossary(
    *,
    brief: dict[str, Any],
    examples_dir: Path,
    rubric_notes: str,
    dry_run: bool,
) -> tuple[dict[str, Any], glossary_rubric.GlossaryLayout]:
    layout = glossary_rubric.get_glossary_layout(examples_dir)
    schema_text = glossary_rubric.glossary_schema_json_text(layout)
    ex_text = serialize_examples_dir(examples_dir)
    if not ex_text.strip():
        logger.warning("No .xlsx examples in %s — using minimal template only.", examples_dir)
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
        logger.info("Dry-run: would call Claude with Glossary schema (%s chars user prompt).", len(user))
        return glossary_rubric.glossary_json_template(layout), layout
    raw = call_claude_json(agent_dir=AGENT_DIR, system_prompt=system, user_prompt=user)
    data = glossary_rubric.normalize_glossary_payload(raw, layout)
    return data, layout


def save_strategy_output(data: dict[str, Any], *, out_path: Path, jira_url: str) -> None:
    wb = strategy_rubric.build_strategy_workbook(data, jira_url=jira_url)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def save_glossary_output(
    data: dict[str, Any],
    layout: glossary_rubric.GlossaryLayout,
    *,
    examples_dir: Path,
    out_path: Path,
    jira_url: str,
) -> None:
    tpl = glossary_rubric.pick_template_path(examples_dir)
    if tpl is None:
        wb = glossary_rubric.create_minimal_glossary_workbook()
    else:
        wb = load_workbook(tpl)
    glossary_rubric.fill_glossary_workbook(wb, data, jira_url=jira_url, layout=layout)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    wb.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="OT Copywriting Agent")
    parser.add_argument("--rubric", required=True, help='Rubric name, e.g. "Strategy" or "First steps in trading"')
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--all", action="store_true", help="Process all .xlsx in the rubric Tasks folder")
    g.add_argument("--input", type=str, help="Single task file relative to project root or absolute")
    parser.add_argument("--dry-run", action="store_true", help="No API calls; validate paths only")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )

    rubric = args.rubric.strip()
    if rubric not in IMPLEMENTED:
        if rubric == "UGC":
            print(
                "UGC rubric is scaffolded only: add schema in rubrics/ugc.py after reviewing UGC Examples."
            )
            return 2
        if rubric == "First steps in trading":
            print(
                "First steps in trading is scaffolded only: add schema in rubrics/first_steps.py "
                "after reviewing First steps in trading Examples."
            )
            return 2
        print(
            f"Rubric '{rubric}' is not implemented yet. Implemented: {', '.join(sorted(IMPLEMENTED))}."
        )
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
                    brief=brief,
                    examples_dir=examples_dir,
                    rubric_notes=notes,
                    dry_run=args.dry_run,
                )
                if not args.dry_run:
                    save_strategy_output(data, out_path=out_path, jira_url=jira_url)
            else:
                data, layout = process_glossary(
                    brief=brief,
                    examples_dir=examples_dir,
                    rubric_notes=notes,
                    dry_run=args.dry_run,
                )
                if not args.dry_run:
                    save_glossary_output(
                        data,
                        layout,
                        examples_dir=examples_dir,
                        out_path=out_path,
                        jira_url=jira_url,
                    )
        except RuntimeError as e:
            if "401" in str(e):
                return 1
            raise
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
