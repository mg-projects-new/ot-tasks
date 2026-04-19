"""Strategy carousel rubric: English-only. Cards 1–7, Card 8 IG/FB, IG/FB/TG posts."""

from __future__ import annotations

import copy
import json
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from excel_io.fill_template import apply_sheet_defaults


# Row labels in output order
ROW_ORDER: list[tuple[str, str]] = [
    ("card_1", "Card 1"),
    ("card_2", "Card 2"),
    ("card_3", "Card 3"),
    ("card_4", "Card 4"),
    ("card_5", "Card 5"),
    ("card_6", "Card 6"),
    ("card_7", "Card 7"),
    ("card8_ig", "Card 8 (IG)"),
    ("card8_fb", "Card 8 (FB)"),
    ("ig_post", "IG Post"),
    ("fb_post", "FB Post"),
    ("tg_post", "TG Post"),
]


def strategy_json_template() -> dict[str, Any]:
    """Flat English-only schema: one string per row."""
    return {key: "" for key, _ in ROW_ORDER}


def strategy_schema_json_text() -> str:
    return json.dumps(strategy_json_template(), ensure_ascii=False, indent=2)


def fill_strategy_workbook(wb: Workbook, data: dict[str, Any], *, jira_url: str) -> None:
    """Populate the Strategy workbook with 3 columns: A=Label, B=EN, C=Chars."""
    if "Лист1" in wb.sheetnames:
        ws = wb["Лист1"]
    else:
        ws = wb.active
        ws.title = "Лист1"

    font = Font(name="Arial", size=10)
    bold = Font(name="Arial", size=10, bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    # Row 1: Jira link
    ws["A1"] = "Jira Task Link →"
    ws["A1"].font = font
    ws["A1"].alignment = wrap
    ws["B1"] = jira_url
    ws["B1"].font = font

    # Row 2: headers
    ws["A2"] = None
    ws["B2"] = "EN"
    ws["C2"] = "Chars"
    ws["B2"].font = bold
    ws["C2"].font = bold
    ws["B2"].alignment = wrap
    ws["C2"].alignment = wrap

    # Row 3+: label + English copy + =LEN()
    for i, (key, display) in enumerate(ROW_ORDER, start=3):
        ws.cell(i, 1, display)
        ws.cell(i, 1).font = font
        ws.cell(i, 1).alignment = wrap
        val = data.get(key, "")
        ws.cell(i, 2, "" if val is None else str(val))
        ws.cell(i, 2).font = font
        ws.cell(i, 2).alignment = wrap
        ws.cell(i, 3, f"=LEN(B{i})")
        ws.cell(i, 3).font = font
        ws.cell(i, 3).alignment = wrap

    # Column widths
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 8

    apply_sheet_defaults(ws, freeze_row=2)


def build_strategy_workbook(data: dict[str, Any], *, jira_url: str) -> Workbook:
    wb = Workbook()
    fill_strategy_workbook(wb, data, jira_url=jira_url)
    return wb


def normalize_strategy_payload(raw: dict[str, Any]) -> dict[str, Any]:
    """Merge model output into the flat template, handling legacy nested shapes."""
    out = strategy_json_template()
    if not isinstance(raw, dict):
        return out

    # Flat schema: {"card_1": "...", "card_2": "..."}
    for key in out:
        if key in raw and isinstance(raw[key], (str, int, float)):
            out[key] = str(raw[key])

    # Back-compat: if model returns {"cards": {"1": "..."}}
    if "cards" in raw and isinstance(raw["cards"], dict):
        for i in range(1, 8):
            v = raw["cards"].get(str(i))
            if isinstance(v, str):
                out[f"card_{i}"] = v
            elif isinstance(v, dict):
                out[f"card_{i}"] = str(v.get("en", "") or "")

    # Back-compat: nested locale dicts for post rows
    for key in ("card8_ig", "card8_fb", "ig_post", "fb_post", "tg_post"):
        if key in raw and isinstance(raw[key], dict):
            out[key] = str(raw[key].get("en", "") or "")

    return out