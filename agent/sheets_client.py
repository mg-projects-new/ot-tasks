"""Google Drive + Sheets client. Handles OAuth, Sheet creation, cell population."""

from __future__ import annotations

import logging
import os
import string
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",   # create/modify files we make
    "https://www.googleapis.com/auth/spreadsheets",  # read/write our Sheets
    "https://www.googleapis.com/auth/documents.readonly",  # read expert-source docs
]

CREDENTIALS_FILE = "google_credentials.json"
TOKEN_FILE = "google_token.json"


# ---------- Auth ----------

def _get_credentials(agent_dir: Path) -> Credentials:
    """OAuth flow: reuse saved token, refresh, or do first-run browser auth."""
    creds_path = agent_dir / CREDENTIALS_FILE
    token_path = agent_dir / TOKEN_FILE

    if not creds_path.exists():
        raise RuntimeError(
            f"Missing {CREDENTIALS_FILE} in agent/. "
            "Download OAuth Desktop credentials from Google Cloud Console and place them there."
        )

    creds: Credentials | None = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return creds


def _services(agent_dir: Path) -> tuple[Any, Any]:
    creds = _get_credentials(agent_dir)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive, sheets


# ---------- Drive helpers ----------

def get_output_folder_id() -> str:
    from dotenv import load_dotenv
    load_dotenv()
    fid = (os.environ.get("GDRIVE_OUTPUT_FOLDER_ID") or "").strip()
    if not fid:
        raise RuntimeError(
            "GDRIVE_OUTPUT_FOLDER_ID not set. Add it to agent/.env with your target Drive folder ID."
        )
    return fid


def _col_letter(idx_1: int) -> str:
    """1-indexed column number → A1 letter (1→A, 27→AA)."""
    s = ""
    n = idx_1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------- Sheet creation ----------

def create_sheet_in_folder(
    agent_dir: Path, *, title: str, folder_id: str,
) -> tuple[str, str]:
    """Create a new Google Sheet inside the given Drive folder.
    Returns (spreadsheet_id, web_url)."""
    drive, sheets = _services(agent_dir)

    # 1. Create the spreadsheet
    body = {"properties": {"title": title}}
    ss = sheets.spreadsheets().create(body=body, fields="spreadsheetId,spreadsheetUrl").execute()
    ss_id = ss["spreadsheetId"]
    ss_url = ss["spreadsheetUrl"]

    # 2. Move it into our target folder (new files land in root by default)
    file = drive.files().get(fileId=ss_id, fields="parents").execute()
    previous = ",".join(file.get("parents") or [])
    drive.files().update(
        fileId=ss_id,
        addParents=folder_id,
        removeParents=previous,
        fields="id,parents",
    ).execute()

    return ss_id, ss_url


def rename_default_sheet(sheets_api, *, spreadsheet_id: str, new_name: str) -> int:
    """Rename the default 'Sheet1' tab. Returns the sheet_id (tab-level numeric id)."""
    meta = sheets_api.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    first = meta["sheets"][0]
    sheet_id = first["properties"]["sheetId"]
    requests_body = {
        "requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "title": new_name},
                "fields": "title",
            }
        }]
    }
    sheets_api.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=requests_body).execute()
    return sheet_id


def add_sheet_tab(sheets_api, *, spreadsheet_id: str, title: str) -> int:
    """Add a new tab, return its sheetId."""
    body = {"requests": [{"addSheet": {"properties": {"title": title[:100]}}}]}
    resp = sheets_api.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def write_values(
    sheets_api, *, spreadsheet_id: str, tab_name: str,
    a1_range: str, values: list[list[Any]], input_option: str = "USER_ENTERED",
) -> None:
    """Write a 2D array to a range. USER_ENTERED makes formulas work."""
    sheets_api.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{tab_name}!{a1_range}",
        valueInputOption=input_option,
        body={"values": values},
    ).execute()


def format_header_and_widths(
    sheets_api, *, spreadsheet_id: str, sheet_id: int,
    header_row: int = 2,
    col_widths: list[tuple[int, int]] | None = None,  # [(col_index_0, width_px), ...]
) -> None:
    """Bold + freeze the header row, optionally set column widths, apply wrap."""
    requests = []

    # Freeze rows above header_row
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": header_row}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # Bold header row
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": header_row - 1,
                "endRowIndex": header_row,
            },
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat.bold",
        }
    })

    # Wrap all cells
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id},
            "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"}},
            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",
        }
    })

    if col_widths:
        for col_index_0, width_px in col_widths:
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": col_index_0,
                        "endIndex": col_index_0 + 1,
                    },
                    "properties": {"pixelSize": width_px},
                    "fields": "pixelSize",
                }
            })

    if requests:
        sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()


def delete_default_sheet_if_present(sheets_api, *, spreadsheet_id: str, keep_titles: set[str]) -> None:
    """After building real tabs, delete any leftover 'Sheet1' that's empty."""
    meta = sheets_api.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in meta.get("sheets", []):
        props = s["properties"]
        if props["title"] in keep_titles:
            continue
        if props["title"] == "Sheet1":
            sheets_api.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"deleteSheet": {"sheetId": props["sheetId"]}}]},
            ).execute()
            return


# ---------- Per-rubric Sheet builders ----------

def build_glossary_sheet(
    agent_dir: Path, *, title: str, jira_url: str, data: dict[str, str],
    row_order: list[str], row_display: dict[str, str], folder_id: str,
) -> tuple[str, str]:
    """Create a Glossary Google Sheet: A=Label, B=EN, C=Chars (with =LEN formulas)."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)
    tab = "Glossary"
    sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=tab)

    # Header rows + body
    values: list[list[Any]] = []
    values.append(["Jira Task Link →", jira_url])      # Row 1
    values.append([None, "EN", "Chars"])               # Row 2

    for key in row_order:
        r = len(values) + 1  # Row number when appended
        cell_val = data.get(key, "") or ""
        values.append([row_display[key], cell_val, f"=LEN(B{r})"])

    a1_end = _col_letter(3)
    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=tab,
        a1_range=f"A1:{a1_end}{len(values)}", values=values,
    )

    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=sheet_id, header_row=2,
        col_widths=[(0, 160), (1, 560), (2, 70)],
    )
    return ss_id, ss_url


def build_strategy_sheet(
    agent_dir: Path, *, title: str, jira_url: str, data: dict[str, str],
    row_order: list[tuple[str, str]], folder_id: str,
) -> tuple[str, str]:
    """Create a Strategy Google Sheet: A=Label, B=EN, C=Chars."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)
    tab = "Strategy"
    sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=tab)

    values: list[list[Any]] = []
    values.append(["Jira Task Link →", jira_url])
    values.append([None, "EN", "Chars"])

    for key, display in row_order:
        r = len(values) + 1
        cell_val = data.get(key, "") or ""
        values.append([display, cell_val, f"=LEN(B{r})"])

    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=tab,
        a1_range=f"A1:C{len(values)}", values=values,
    )

    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=sheet_id, header_row=2,
        col_widths=[(0, 160), (1, 560), (2, 70)],
    )
    return ss_id, ss_url


def build_ugc_sheet(
    agent_dir: Path, *, title: str, jira_url: str, videos: list[dict[str, Any]], folder_id: str,
) -> tuple[str, str]:
    """Create a UGC Google Sheet with one tab per video."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)

    # Rename the default sheet to match the first video
    if not videos:
        # Edge case: no videos to render
        rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name="Video 1")
        return ss_id, ss_url

    first = videos[0]
    first_tab_name = _ugc_tab_name(first)
    first_sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=first_tab_name)
    _populate_ugc_tab(sheets, ss_id, first_tab_name, first_sheet_id, first, jira_url)

    for v in videos[1:]:
        tab = _ugc_tab_name(v)
        sid = add_sheet_tab(sheets, spreadsheet_id=ss_id, title=tab)
        _populate_ugc_tab(sheets, ss_id, tab, sid, v, jira_url)

    return ss_id, ss_url


def _ugc_tab_name(v: dict[str, Any]) -> str:
    n = int(v.get("video_number") or 1)
    script = str(v.get("script_id") or "").strip()
    if script:
        return f"Video {n} (Script {script})"
    return f"Video {n}"


def _populate_ugc_tab(
    sheets_api, spreadsheet_id: str, tab: str, sheet_id: int,
    video: dict[str, Any], jira_url: str,
) -> None:
    cover = str(video.get("cover") or "")
    caption = str(video.get("caption") or "")
    notes = (
        "1. Don't translate the Image Title row.  "
        "2. The entire text of the post should be written in one cell.  "
        "3. Separate paragraphs by a blank line within the same cell."
    )

    values = [
        ["Jira Task Link →", jira_url],                             # Row 1
        [None, "EN", None, "Chars"],                                 # Row 2
        ["TikTok video - Cover", cover, None, "=LEN(B3)"],           # Row 3
        ["TikTok video1", caption, None, "=LEN(B4)"],                # Row 4
        [None], [None], [None], [None], [None], [None], [None], [None],  # Rows 5-12
        [notes],                                                     # Row 13
    ]
    write_values(
        sheets_api, spreadsheet_id=spreadsheet_id, tab_name=tab,
        a1_range=f"A1:D{len(values)}", values=values,
    )
    format_header_and_widths(
        sheets_api, spreadsheet_id=spreadsheet_id, sheet_id=sheet_id, header_row=2,
        col_widths=[(0, 170), (1, 400), (2, 50), (3, 70)],
    )


def build_first_steps_sheet(
    agent_dir: Path, *, title: str, jira_url: str, data: dict[str, str],
    row_order: list[tuple[str, str]], folder_id: str,
) -> tuple[str, str]:
    """Create a First steps in trading Google Sheet: A=Label, B=EN, C=Chars."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)
    tab = "First steps"
    sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=tab)

    values: list[list[Any]] = []
    values.append(["Jira Task Link \u2192", jira_url])
    values.append([None, "EN", "Chars"])

    for key, display in row_order:
        r = len(values) + 1
        cell_val = data.get(key, "") or ""
        values.append([display, cell_val, f"=LEN(B{r})"])

    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=tab,
        a1_range=f"A1:C{len(values)}", values=values,
    )

    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=sheet_id, header_row=2,
        col_widths=[(0, 160), (1, 560), (2, 70)],
    )
    return ss_id, ss_url



def build_ongoing_p1_sheet(
    agent_dir: Path, *, title: str, jira_url: str, data: dict[str, Any], folder_id: str,
) -> tuple[str, str]:
    """Create a native Google Sheet for Ongoing P1: single row, single tab."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)
    tab = "Economic calendar"
    sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=tab)

    image_title = data.get("image_title", "") or ""
    notes = (
        "1. Translate the Image Title row. "
        "2. The entire text of the post should be written in one cell. "
        "3. Separate paragraphs by one paragraph break (Alt/Option + Enter) as you'd do that in a doc. "
        "4. SM posts can't have text styling. So please don't apply any."
    )

    values = [
        ["Jira Task Link →", jira_url],
        [None, "EN", "Chars"],
        ["Image Title  (AR, TH, FA)", image_title, "=LEN(B3)"],
        [None],
        [notes],
    ]
    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=tab,
        a1_range=f"A1:C{len(values)}", values=values,
    )
    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=sheet_id, header_row=2,
        col_widths=[(0, 200), (1, 700), (2, 70)],
    )
    return ss_id, ss_url


def build_ongoing_p2_sheet(
    agent_dir: Path, *, title: str, jira_url: str, data: dict[str, Any], folder_id: str,
) -> tuple[str, str]:
    """Create a native Google Sheet for Ongoing P2: two tabs (Trading Signal + Asset of the day)."""
    ss_id, ss_url = create_sheet_in_folder(agent_dir, title=title, folder_id=folder_id)
    _, sheets = _services(agent_dir)

    # Tab 1: Trading Signal
    ts_data = data.get("trading_signal") or {}
    ts_tab = "Trading Signal"
    ts_sheet_id = rename_default_sheet(sheets, spreadsheet_id=ss_id, new_name=ts_tab)

    ts_rows = [
        ("Image Title", ts_data.get("image_title", "")),
        ("TG (max 1024)", ts_data.get("tg", "")),
        ("Button", ts_data.get("button", "")),
        ("TW", ts_data.get("tw", "")),
        ("Poll Option 1", ts_data.get("poll_option_1", "")),
        ("Poll Option 2", ts_data.get("poll_option_2", "")),
    ]

    ts_values = [
        ["Jira Task Link →", jira_url],
        [None, "EN", "Chars"],
    ]
    for i, (label, val) in enumerate(ts_rows, start=3):
        ts_values.append([label, val or "", f"=LEN(B{i})"])

    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=ts_tab,
        a1_range=f"A1:C{len(ts_values)}", values=ts_values,
    )
    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=ts_sheet_id, header_row=2,
        col_widths=[(0, 160), (1, 560), (2, 70)],
    )

    # Tab 2: Asset of the day
    aotd_data = data.get("asset_of_the_day") or {}
    aotd_tab = "Asset of the day"
    aotd_sheet_id = add_sheet_tab(sheets, spreadsheet_id=ss_id, title=aotd_tab)

    aotd_rows = [
        ("Image Title (TG, FB)", aotd_data.get("image_title", "")),
        ("TG  (max 1024)", aotd_data.get("tg", "")),
        ("Button [platform link]", aotd_data.get("button", "")),
        ("FB post", aotd_data.get("fb_post", "")),
    ]

    aotd_values = [
        ["Jira Task Link →", jira_url],
        [None, "EN", "Chars"],
        [None, "Asset of the week"],
    ]
    for i, (label, val) in enumerate(aotd_rows, start=4):
        aotd_values.append([label, val or "", f"=LEN(B{i})"])

    write_values(
        sheets, spreadsheet_id=ss_id, tab_name=aotd_tab,
        a1_range=f"A1:C{len(aotd_values)}", values=aotd_values,
    )
    format_header_and_widths(
        sheets, spreadsheet_id=ss_id, sheet_id=aotd_sheet_id, header_row=2,
        col_widths=[(0, 180), (1, 560), (2, 70)],
    )

    return ss_id, ss_url
