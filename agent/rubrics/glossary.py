"""Glossary rubric: English-only output. Fixed 3-column layout (Label | EN | Chars)."""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from excel_io.fill_template import apply_sheet_defaults


# Row labels in the order they appear in the output xlsx
ROW_ORDER = ["image_title", "ig", "fb", "tg", "button"]
ROW_DISPLAY = {
    "image_title": "Image Title",
    "ig": "IG",
    "fb": "FB",
    "tg": "TG Post",
    "button": "Button",
}


@dataclass
class GlossaryLayout:
    """Kept as a dataclass for back-compat with run.py, but English-only now."""
    sheet_name: str = "Лист1"
    row_labels: dict[str, int] = field(default_factory=dict)


def glossary_json_template() -> dict[str, Any]:
    """Empty template for the prompt schema — flat English-only strings."""
    return {label: "" for label in ROW_ORDER}


def glossary_schema_json_text(layout: GlossaryLayout | None = None) -> str:
    return json.dumps(glossary_json_template(), ensure_ascii=False, indent=2)


def get_glossary_layout(examples_dir: Path) -> GlossaryLayout:
    """Layout is fixed now; returned for back-compat with run.py."""
    layout = GlossaryLayout()
    for i, key in enumerate(ROW_ORDER, start=3):
        layout.row_labels[key] = i
    return layout


def pick_template_path(examples_dir: Path) -> Path | None:
    """Not used in English-only mode — we build a fresh workbook every time."""
    return None


def create_minimal_glossary_workbook() -> Workbook:
    """Build the English-only Glossary sheet from scratch: A=Label, B=EN, C=Chars."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    font = Font(name="Arial", size=10)
    wrap = Alignment(wrap_text=True, vertical="top")

    # Row 1: Jira link
    ws["A1"] = "Jira Task Link →"
    ws["A1"].font = font
    ws["A1"].alignment = wrap
    ws["B1"] = ""
    ws["B1"].font = font

    # Row 2: column headers
    ws["A2"] = None
    ws["B2"] = "EN"
    ws["C2"] = "Chars"
    for c in ("B2", "C2"):
        ws[c].font = Font(name="Arial", size=10, bold=True)
        ws[c].alignment = wrap

    # Row 3+: label rows
    for i, key in enumerate(ROW_ORDER, start=3):
        ws.cell(i, 1, ROW_DISPLAY[key])
        ws.cell(i, 1).font = font
        ws.cell(i, 1).alignment = wrap
        ws.cell(i, 2, "")
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
    return wb


def load_layout_from_workbook(wb: Workbook) -> GlossaryLayout:
    """Back-compat no-op: return the fixed English layout."""
    layout = GlossaryLayout(sheet_name=wb.active.title)
    for i, key in enumerate(ROW_ORDER, start=3):
        layout.row_labels[key] = i
    return layout


def fill_glossary_workbook(
    wb: Workbook,
    data: dict[str, Any],
    *,
    jira_url: str,
    layout: GlossaryLayout,
) -> None:
    """Write English copy into column B of each label row."""
    ws = wb[layout.sheet_name] if layout.sheet_name in wb.sheetnames else wb.active
    font = Font(name="Arial", size=10)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws["B1"] = jira_url
    ws["B1"].font = font

    for key in ROW_ORDER:
        row_idx = layout.row_labels.get(key)
        if not row_idx:
            continue
        val = data.get(key, "")
        text = "" if val is None else str(val)
        ws.cell(row_idx, 2, text)
        ws.cell(row_idx, 2).font = font
        ws.cell(row_idx, 2).alignment = wrap
        ws.cell(row_idx, 3, f"=LEN(B{row_idx})")
        ws.cell(row_idx, 3).font = font
        ws.cell(row_idx, 3).alignment = wrap

    apply_sheet_defaults(ws, freeze_row=2)


def normalize_glossary_payload(raw: dict[str, Any], layout: GlossaryLayout | None = None) -> dict[str, Any]:
    """Pull top-level string values per row key, ignoring any locale nesting if the model returns it."""
    out = {key: "" for key in ROW_ORDER}
    if not isinstance(raw, dict):
        return out
    for key in ROW_ORDER:
        if key not in raw:
            continue
        v = raw[key]
        if isinstance(v, str):
            out[key] = v
        elif isinstance(v, dict):
            # If model produces nested locale dict anyway, take "en"
            out[key] = str(v.get("en", "") or "")
        elif v is not None:
            out[key] = str(v)
    return out