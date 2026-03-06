#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ============================================================
# Canonical "Daily list PC version" — Dark Blue Tag Text
#
# Changes in this version:
# - Prompts for MONTH first (remembers last used month in last_used.json).
# - Then asks for SERVICE DAY (number only).
# - Dailylist sheet resolution accepts BOTH full and short month names:
#     "Dailylist October" OR "dailylist oct" (case-insensitive, extra spaces tolerated).
# - All other behavior unchanged.
# ============================================================

PAGE3_COLS = 4
PAGE3_ROWS = 56

# Manual squeeze for Page 1 (fit to ONE page)
PAGE1_FORCE_ONE_PAGE = True

# Per-page orientation: 'p' = portrait, 'l' = landscape
PAGE1_ORIENT = 'l'   # client list (C..M)
PAGE2_ORIENT = 'p'   # meal list (P..Z)
PAGE3_ORIENT = 'p'   # mise list (auto-find)
PAGE4_ORIENT = 'p'   # delivery list (Delivery sheet block)

# Delivery page configuration
DELIVERY_ENABLED = True
DELIVERY_BLOCK_WIDTH = 4
DELIVERY_DATE_ROW = 1
DELIVERY_DATA_START_ROW = 2

# Global export margin (in inches) used for Sheets PDF export.
EXPORT_MARGIN_INCH = 0.12  # reduced from 0.25 to use more printable area

# --------- Tag text color (single-color change) ----------
TAG_TEXT_COLOR = "#000000"   # dark blue for all tag text
# ============================================================

import os
import re
import json
import time
import sys
from io import BytesIO, StringIO
from datetime import datetime, timedelta, date
from typing import Tuple, Optional, Dict, Any, List
import calendar
from google.oauth2 import service_account

from google.auth.transport.requests import AuthorizedSession, Request

try:
    from PyPDF2 import PdfMerger, PdfReader
except ImportError as e:
    raise SystemExit("PyPDF2 is required. Install it with: pip install PyPDF2") from e

# ---------- Config ----------
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1CsT6_oYsFjgQQ73pt1Bl1cuXuzKY8JnOTB3E4bDkTiA/edit?usp=sharing'
LAST_USED_FILE = 'last_used.json'

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

# ============================================================
# Branded single-line 0–100 progress bar (centered text reveal)
# ============================================================

BRAND_TEXT = "Making Friska Wellness List"

def _bar(pct: int):
    pct = max(0, min(100, int(pct)))
    width = 30  # inside the [ ... ] area

    filled = int(width * pct / 100)

    content = [" "] * width
    for i in range(filled):
        content[i] = "="

    blen = len(BRAND_TEXT)
    start = max(0, (width - blen) // 2)
    end = min(width, start + blen)

    overlap = max(0, min(filled, end) - start)
    for i in range(overlap):
        idx = start + i
        if 0 <= idx < width:
            content[idx] = BRAND_TEXT[i]

    line = f"\r[{''.join(content)}] {pct}%"
    sys.stdout.write(line)
    sys.stdout.flush()
    if pct == 100:
        sys.stdout.write("\n")
        sys.stdout.flush()

# ---------- last_used (month + date memory; backward compatible) ----------
def load_last_used() -> Dict[str, Any]:
    """
    Returns dict with optional keys:
      - "month": MonthName (e.g., "November")
      - "date":  last date string you entered earlier (e.g., "01-Nov-25")
    Backward compatible if only "date" exists.
    """
    if os.path.exists(LAST_USED_FILE):
        try:
            with open(LAST_USED_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {"month": "", "date": ""}

def save_last_used(month_name: str = "", date_in: str = "") -> None:
    """
    Save last used month and/or date. Only overwrites fields provided.
    """
    data = load_last_used()
    if month_name:
        data["month"] = month_name
    if date_in:
        data["date"] = date_in
    with open(LAST_USED_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_creds():

    try:
        import streamlit as st
        import base64
        import json

        sa_json = base64.b64decode(
            st.secrets["google"]["service_account_base64"]
        ).decode("utf-8")

        sa_info = json.loads(sa_json)

    except Exception:
        with open("service_account.json", "r", encoding="utf-8") as f:
            sa_info = json.load(f)

    creds = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=SCOPES
    )

    return creds

def get_spreadsheet_id_from_url(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        raise ValueError("Could not parse spreadsheetId from URL.")
    return m.group(1)

# ---------- HTTP helper ----------
def http_get_with_retry(authed: AuthorizedSession, url: str, params: Optional[Dict[str, str]] = None,
                        max_attempts=4, backoff=1.6):
    last_exc: Optional[Exception] = None
    for attempt in range(1, max_attempts + 1):
        try:
            resp = authed.get(url, params=params, timeout=30)
            if 200 <= resp.status_code < 300:
                return resp
            if 500 <= resp.status_code < 600:
                raise RuntimeError(f"HTTP {resp.status_code}")
            resp.raise_for_status()
        except Exception as e:
            last_exc = e
            if attempt == max_attempts:
                break
            time.sleep(backoff ** attempt)
    if last_exc:
        raise last_exc
    raise RuntimeError("Request failed")

# ---------- Sheets metadata ----------
def get_spreadsheet_metadata(authed: AuthorizedSession, spreadsheet_id: str) -> Dict[str, Any]:
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
    resp = http_get_with_retry(authed, meta_url)
    return resp.json()

def get_ordered_sheet_titles(authed: AuthorizedSession, spreadsheet_id: str) -> List[str]:
    data = get_spreadsheet_metadata(authed, spreadsheet_id)
    ordered = sorted(
        [(sh["properties"]["index"], sh["properties"]["title"]) for sh in data.get("sheets", [])],
        key=lambda x: x[0]
    )
    return [t for _, t in ordered]

def _norm_spaces_ci(s: str) -> str:
    """Lowercase + collapse whitespace for robust comparisons."""
    return re.sub(r"\s+", " ", (s or "").strip()).lower()

def resolve_sheet_title_ci(authed: AuthorizedSession, spreadsheet_id: str, desired_name: str) -> str:
    """
    Case-insensitive match that also normalizes internal whitespace:
    - Collapses multiple spaces to a single space
    - Trims leading/trailing spaces
    Still requires the same words in the same order.
    """
    data = get_spreadsheet_metadata(authed, spreadsheet_id)
    desired_norm = _norm_spaces_ci(desired_name)

    # 1) strict equal after normalization
    for sh in data.get("sheets", []):
        title = sh.get("properties", {}).get("title", "")
        if _norm_spaces_ci(title) == desired_norm:
            return title

    # 2) fallback: substring after normalization
    for sh in data.get("sheets", []):
        title = sh.get("properties", {}).get("title", "")
        if desired_norm in _norm_spaces_ci(title):
            return title

    available = [sh.get("properties", {}).get("title", "") for sh in data.get("sheets", [])]
    raise SystemExit(
        f"Sheet/tab '{desired_name}' not found.\nAvailable tabs:\n- " + "\n- ".join(available)
    )

def resolve_dailylist_for_month(authed: AuthorizedSession, spreadsheet_id: str, month_full_name: str) -> str:
    """
    Resolve the Dailylist sheet for a given month, accepting:
      - 'Dailylist <FullMonth>'   e.g., 'Dailylist October'
      - 'Dailylist <MonAbbr>'     e.g., 'Dailylist Oct'
    Matching is case-insensitive and collapses extra spaces.
    """
    month_full = month_full_name.strip()
    month_abbr = month_full[:3]  # Oct, Nov, etc.

    want_full = _norm_spaces_ci(f"Dailylist {month_full}")
    want_abbr = _norm_spaces_ci(f"Dailylist {month_abbr}")

    meta = get_spreadsheet_metadata(authed, spreadsheet_id)
    candidates = [sh.get("properties", {}).get("title", "") for sh in meta.get("sheets", [])]

    # 1) exact (after whitespace normalization)
    for t in candidates:
        tn = _norm_spaces_ci(t)
        if tn == want_full or tn == want_abbr:
            return t

    # 2) startswith 'dailylist ' and the rest equals month token
    for t in candidates:
        tn = _norm_spaces_ci(t)
        if tn.startswith("dailylist "):
            tail = tn[len("dailylist "):].strip()
            if tail == month_full.lower() or tail == month_abbr.lower():
                return t

    # 3) Helpful error with available Dailylist tabs
    daily_tabs = [c for c in candidates if _norm_spaces_ci(c).startswith("dailylist ")]
    msg = "Sheet/tab 'Dailylist {0}' (or 'Dailylist {1}') not found.\nAvailable Dailylist-like tabs:\n".format(
        month_full, month_abbr
    )
    if daily_tabs:
        msg += "- " + "\n- ".join(daily_tabs)
    else:
        msg += "(none)"
    raise SystemExit(msg)

def get_gid_for_sheet(title: str, spreadsheet_id: str, authed: AuthorizedSession) -> int:
    data = get_spreadsheet_metadata(authed, spreadsheet_id)
    for sh in data.get("sheets", []):
        props = sh.get("properties", {})
        if props.get("title") == title:
            return int(props.get("sheetId"))
    raise ValueError(f"Sheet/tab named '{title}' not found.")

# ---------- Export EXACT range (returns PDF bytes) ----------
def export_range_pdf_bytes(spreadsheet_id: str, gid: int, r1: int, c1: int, r2: int, c2: int,
                           authed: AuthorizedSession, portrait: bool, fit_page: bool) -> bytes:
    base = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
    margin = str(EXPORT_MARGIN_INCH)
    params = {
        "format": "pdf",
        "gid": str(gid),
        "r1": str(r1),
        "c1": str(c1),
        "r2": str(r2),
        "c2": str(c2),
        "size": "A4",
        "portrait": "true" if portrait else "false",
        "sheetnames": "false",
        "printtitle": "false",
        "pagenum": "UNDEFINED",
        "gridlines": "false",
        "fzr": "false",
        "hf": "0",
        "top_margin": margin,
        "bottom_margin": margin,
        "left_margin": margin,
        "right_margin": margin,
    }
    if fit_page:
        params["scale"] = "4"      # Fit to page (force onto one page)
    else:
        params["fitw"] = "true"    # Fit to width (may spill to multiple pages)

    resp = http_get_with_retry(authed, base, params=params)
    resp.raise_for_status()
    return resp.content

# ---------- Values ----------
from urllib.parse import quote

def get_values_safe(authed: AuthorizedSession, spreadsheet_id: str, a1_range: str) -> List[List[str]]:
    encoded_range = quote(a1_range, safe="")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{encoded_range}"
    params = {
        "majorDimension": "ROWS",
        "valueRenderOption": "UNFORMATTED_VALUE",
    }
    resp = http_get_with_retry(authed, url, params=params)
    data = resp.json()
    return data.get("values", [])

def get_values_block(authed: AuthorizedSession, spreadsheet_id: str,
                     sheet_name: str, col_start: str, col_end: str,
                     row_start: int, row_end: int) -> List[List[str]]:
    if row_end < row_start:
        return []
    a1 = f"{sheet_name}!{col_start}{row_start}:{col_end}{row_end}"
    return get_values_safe(authed, spreadsheet_id, a1)

def get_sheet_values_full(authed: AuthorizedSession, spreadsheet_id: str, sheet_name: str,
                          max_cols: int = 200, max_rows: int = 10000) -> List[List[str]]:
    end_col_letter = idx_to_col_letter(max_cols - 1)
    a1 = f"{sheet_name}!A1:{end_col_letter}{max_rows}"
    return get_values_safe(authed, spreadsheet_id, a1)

def row_has_any_value(row_values: List[str]) -> bool:
    if not row_values:
        return False
    for v in row_values:
        if isinstance(v, (int, float)):
            return True
        if str(v).strip() != "":
            return True
    return False

# ---------- Date normalization ----------
MONTHS = {
    'jan':1,'january':1,'feb':2,'february':2,'mar':3,'march':3,'apr':4,'april':4,
    'may':5,'jun':6,'june':6,'jul':7,'july':7,'aug':8,'august':8,'sep':9,'sept':9,'september':9,
    'oct':10,'october':10,'nov':11,'november':11,'dec':12,'december':12,
}

def normalize_input_date(user_input: str) -> str:
    s = user_input.strip()
    if not s:
        raise ValueError("Empty date.")

    # If the user enters only a day number, assume current month & year
    if re.fullmatch(r"\d{1,2}", s):
        today = datetime.now()
        day = int(s)
        dtv = datetime(today.year, today.month, day)  # raises if invalid day
        return dtv.strftime("%d-%b-%y")

    candidates = [
        "%d-%b-%y", "%d-%b-%Y",
        "%d %b %y", "%d %b %Y",
        "%d %B %y", "%d %B %Y",
        "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y",
        "%m/%d/%Y", "%m/%d/%y",
        "%b %d %Y", "%B %d %Y", "%b %d, %Y", "%B %d, %Y",
        "%d-%m-%Y", "%d-%m-%y"
    ]
    for fmt in candidates:
        try:
            dtv = datetime.strptime(s, fmt)
            return dtv.strftime("%d-%b-%y")
        except Exception:
            pass

    m = re.match(r"^\s*(\d{1,2})\s*[-/\s]?\s*([A-Za-z]+)\s*(\d{2,4})?\s*$", s)
    if not m:
        m = re.match(r"^\s*([A-Za-z]+)\s*(\d{1,2})\s*(\d{2,4})?\s*$", s)
        if m:
            if m.group(3):
                s2 = f"{m.group(2)} {m.group(1)} {m.group(3)}"
            else:
                s2 = f"{m.group(2)} {m.group(1)}"
            return normalize_input_date(s2)
        raise ValueError("Unrecognized date format.")

    day = int(m.group(1))
    mon_str = m.group(2).lower()
    year = m.group(3)
    if mon_str not in MONTHS:
        raise ValueError("Unknown month name.")
    month = MONTHS[mon_str]
    if year is None:
        year = datetime.now().year
    else:
        year = int(year)
        if year < 100:
            year += 2000
    dtv = datetime(year, month, day)
    return dtv.strftime("%d-%b-%y")

# --- Date parsing helpers ---
WEEKDAY_PREFIX = re.compile(r"^\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s*,\s*", re.I)

def to_dt(val: object) -> Optional[datetime]:
    if isinstance(val, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(val))
        except Exception:
            return None
    s = str(val).strip()
    if not s:
        return None
    s = WEEKDAY_PREFIX.sub("", s)
    try:
        norm = normalize_input_date(s)
        return datetime.strptime(norm, "%d-%b-%y")
    except Exception:
        return None

def as_sheet_date_text(val: object) -> str:
    if isinstance(val, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(val))
            return dt.strftime("%d-%b-%y")
        except Exception:
            return str(val).strip()
    s = str(val).strip()
    if not s:
        return ""
    s_wo_weekday = WEEKDAY_PREFIX.sub("", s)
    try:
        return normalize_input_date(s_wo_weekday)
    except Exception:
        return s

# ---------- Merged-cell helpers ----------
def col_letter_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1

def idx_to_col_letter(idx: int) -> str:
    s = ""
    x = idx + 1
    while x > 0:
        x, r = divmod(x - 1, 26)
        s = chr(65 + r) + s
    return s

def row_is_covered_by_vertical_merge(row_1based: int, col_start_letter: str, col_end_letter: str,
                                     merges: list) -> bool:
    r = row_1based - 1
    c1 = col_letter_to_idx(col_start_letter)
    c2 = col_letter_to_idx(col_end_letter) + 1
    for m in merges:
        sr = m.get("startRowIndex", 0)
        er = m.get("endRowIndex", 0)
        sc = m.get("startColumnIndex", 0)
        ec = m.get("endColumnIndex", 0)
        if er - sr <= 1:
            continue
        if not (sr <= r < er):
            continue
        if not (sc < c2 and ec > c1):
            continue
        return True
    return False

def find_end_by_two_empty_rows(
    authed: AuthorizedSession,
    spreadsheet_id: str,
    sheet_name: str,
    col_start: str,
    col_end: str,
    start_row: int,
    max_row: int = 9999,
    empty_run: int = 2,
    merges: Optional[list] = None
) -> int:
    step = 200
    empties = 0
    last_with_data = start_row - 1
    r = start_row

    while r <= max_row:
        chunk_end = min(r + step - 1, max_row)

        try:
            block = get_values_block(
                authed,
                spreadsheet_id,
                sheet_name,
                col_start,
                col_end,
                r,
                chunk_end
            )
        except Exception:
            # SAFETY FIX:
            # Google Sheets can throw HTTP 400 for invalid / sparse ranges
            # (very common on month-end dates like 30th/31st)
            break

        if not block:
            break

        for i, row_vals in enumerate(block, start=r):
            row_has_data = row_has_any_value(row_vals)

            if not row_has_data and merges:
                if row_is_covered_by_vertical_merge(i, col_start, col_end, merges):
                    row_has_data = True

            if row_has_data:
                last_with_data = i
                empties = 0
            else:
                empties += 1
                if empties >= empty_run:
                    return i - empty_run

        r = chunk_end + 1

    return max(last_with_data, start_row)


# ---------- Start/Next date rows in column C (for a specific sheet) ----------
def find_start_and_next_rows(authed: AuthorizedSession, spreadsheet_id: str, sheet_name: str, date_text: str) -> Tuple[int, Optional[int]]:
    rng = f"{sheet_name}!C1:C9999"
    values = get_values_safe(authed, spreadsheet_id, rng)
    start_row = None
    for i, row in enumerate(values, start=1):
        raw = row[0] if row else ""
        if as_sheet_date_text(raw) == date_text:
            start_row = i
            break
    if start_row is None:
        raise ValueError(f"Date '{date_text}' not found in column C on sheet '{sheet_name}'.")
    next_row = None
    start_dt = datetime.strptime(date_text, "%d-%b-%y")
    threshold = start_dt + timedelta(days=1)
    for i in range(start_row + 1, len(values) + 1):
        raw = values[i - 1][0] if (i - 1) < len(values) and values[i - 1] else ""
        cand_dt = to_dt(raw)
        if cand_dt and cand_dt >= threshold:
            next_row = i
            break
    return start_row, next_row

def find_last_used_row_pz_till_sheet_end(
    authed,
    spreadsheet_id,
    sheet_name,
    start_row
) -> int:
    """
    Scan P–Z till physical sheet end.
    Ignore empty rows.
    Return LAST row having ANY data in P–Z.
    """

    meta = get_spreadsheet_metadata(authed, spreadsheet_id)
    max_row = 9999
    for sh in meta.get("sheets", []):
        if sh.get("properties", {}).get("title") == sheet_name:
            max_row = sh["properties"]["gridProperties"]["rowCount"]
            break

    a1 = f"{sheet_name}!P{start_row}:Z{max_row}"
    values = get_values_safe(authed, spreadsheet_id, a1)

    last_with_data = start_row
    for idx, row in enumerate(values):
        absolute_row = start_row + idx
        if row and any(str(v).strip() for v in row if v is not None):
            last_with_data = absolute_row

    return last_with_data




# ---------- Page 3 auto-find on "MiseList" ----------
def auto_find_page3_range_on_miselist(authed: AuthorizedSession, spreadsheet_id: str, normalized_date: str) -> Tuple[str, str]:
    sheet_title = resolve_sheet_title_ci(authed, spreadsheet_id, "MiseList")
    rows = get_sheet_values_full(authed, spreadsheet_id, sheet_title, max_cols=200, max_rows=10000)

    target_dt = datetime.strptime(normalized_date, "%d-%b-%y").date()
    start_r = start_c = None

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            dtv = to_dt(val)
            if dtv:
                if dtv.date() == target_dt:
                    start_r, start_c = r_idx, c_idx
                    break
            else:
                if as_sheet_date_text(val) == normalized_date:
                    start_r, start_c = r_idx, c_idx
                    break
        if start_r is not None:
            break

    if start_r is None:
        raise SystemExit(f"Could not find date '{normalized_date}' on sheet '{sheet_title}' for page 3.")

    end_r = start_r + PAGE3_ROWS - 1
    end_c = start_c + PAGE3_COLS - 1
    start_col_letter = idx_to_col_letter(start_c - 1)
    end_col_letter   = idx_to_col_letter(end_c - 1)
    a1 = f"{start_col_letter}{start_r}:{end_col_letter}{end_r}"
    return sheet_title, a1

# ---------- Delivery block auto-find ----------
def get_sheet_merges(authed: AuthorizedSession, spreadsheet_id: str, sheet_title: str):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
    params = {"fields": "sheets(properties(sheetId,title),merges)"}
    resp = http_get_with_retry(authed, url, params=params)
    data = resp.json()
    for sh in data.get("sheets", []):
        props = sh.get("properties", {})
        if props.get("title") == sheet_title:
            return sh.get("merges", []) or []
    return []

def auto_find_delivery_block_range(authed: AuthorizedSession, spreadsheet_id: str, normalized_date: str) -> Optional[Tuple[str, str]]:
    if not DELIVERY_ENABLED:
        return None

    sheet_title = resolve_sheet_title_ci(authed, spreadsheet_id, "Delivery")
    max_cols = 200
    end_col = idx_to_col_letter(max_cols - 1)
    header_range = f"{sheet_title}!A{DELIVERY_DATE_ROW}:{end_col}{DELIVERY_DATE_ROW}"
    headers = get_values_safe(authed, spreadsheet_id, header_range)
    if not headers:
        return None
    header_row = headers[0] if headers else []

    target_dt = datetime.strptime(normalized_date, "%d-%b-%y").date()
    found_col_idx = None
    for c_idx, val in enumerate(header_row, start=1):
        dtv = to_dt(val)
        if dtv and dtv.date() == target_dt:
            found_col_idx = c_idx
            break
        if as_sheet_date_text(val) == normalized_date:
            found_col_idx = c_idx
            break

    if found_col_idx is None:
        return None

    start_c = found_col_idx
    end_c = start_c + DELIVERY_BLOCK_WIDTH - 1
    start_col_letter = idx_to_col_letter(start_c - 1)
    end_col_letter = idx_to_col_letter(end_c - 1)

    start_row = DELIVERY_DATA_START_ROW
    end_row = find_end_by_two_empty_rows(authed, spreadsheet_id, sheet_title, start_col_letter, end_col_letter, start_row,
                                         max_row=9999, empty_run=2, merges=get_sheet_merges(authed, spreadsheet_id, sheet_title))
    a1 = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"
    return sheet_title, a1

# ---------- A1 to row/col ----------
def a1_to_rc(a1: str) -> Tuple[int, int, int, int]:
    a1 = a1.strip().upper()
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", a1)
    if not m:
        raise ValueError("Invalid A1 range. Example: C603:Z645")
    def col_to_idx(col: str) -> int:
        n = 0
        for ch in col:
            n = n * 26 + (ord(ch) - 64)
        return n - 1
    c1 = col_to_idx(m.group(1))
    r1 = int(m.group(2)) - 1
    c2 = col_to_idx(m.group(3)) + 1
    r2 = int(m.group(4))
    return r1, c1, r2, c2

# ---------- Merge PDFs from bytes ----------
def merge_pdfs_bytes_to_file(pdf_bytes_list: List[bytes], out_pdf_path: str):
    merger = PdfMerger()
    try:
        for blob in pdf_bytes_list:
            merger.append(BytesIO(blob))
        if os.path.exists(out_pdf_path):
            try:
                os.remove(out_pdf_path)
            except PermissionError:
                merger.close()
                raise SystemExit("Close the output PDF and run again.")
        with open(out_pdf_path, "wb") as f:
            merger.write(f)
        merger.close()
    finally:
        try:
            merger.close()
        except Exception:
            pass

# ============================================================
# TAG GENERATOR (in-memory; fills last page with blanks)
# ============================================================

import csv as _csv, requests as _requests
from io import StringIO as _StringIO
from PIL import Image as _Image, ImageDraw as _ImageDraw, ImageFont as _ImageFont
from reportlab.pdfgen import canvas as _canvas
from reportlab.lib.pagesizes import A4 as _A4
from reportlab.lib.utils import ImageReader as _ImageReader
import os

MEAL_TEMPLATE_FILE = "tag_template.png"
CARRYBAG_TEMPLATE_FILE = "carrybag_tags.png"
DISHES_FILE = "dishes.csv"
BORDER_MARGIN_MM = 7
LINE_SPACING_FACTOR = 0.22
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Column indexes in D..M
COL_CLIENT = 0
COL_BREAKFAST = 1
COL_OPT1 = 2
COL_OPT2 = 3
COL_SNACK = 4
COL_JUICE1 = 5
COL_JUICE2 = 6
COL_REMARKS = 7
COL_TYPE = 8
COL_SLOT = 9

MEAL_FONT_SIZES = {"dish": 85, "meal": 50, "client": 50, "remarks": 70}
CARRYBAG_FONT_SIZES = {"client": 100, "dish": 90, "meal": 60, "remarks": 50}

def _mm_to_px(mm, dpi):
    return int(round(mm * dpi / 25.4))

def _wrap_line(draw, text, font, max_width):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if draw.textlength(test, font=font) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines

def _choose_font_path():
    font_path = os.path.join(BASE_DIR, "mealtag_font.ttf")

    try:
        _ImageFont.truetype(font_path, 12)
        return font_path
    except Exception as e:
        raise RuntimeError(f"Font not found: {font_path}")

def _clean_meal_type(meal_type):
    return re.sub(r"^\s*\d+[\.\-\s]*", "", meal_type)

def _load_dish_map(path):
    m = {}
    with open(path, newline="", encoding="utf-8") as f:
        r = _csv.reader(f)
        next(r, None)
        for row in r:
            if row and len(row) >= 2:
                m[row[0].strip().lower()] = row[1].strip()
    return m

def _fetch_sheet_range_csv(spreadsheet_id: str, sheet_name: str, start_row: int, end_row: int):
    col_range = f"D{start_row}:M{end_row}"
    url = (f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?"
           f"tqx=out:csv&sheet={sheet_name}&range={col_range}")
    resp = _requests.get(url)
    resp.raise_for_status()
    return list(_csv.reader(_StringIO(resp.content.decode("utf-8", errors="replace"))))

def _draw_meal_tag(template_img, texts, font_path, skip_all=False, skip_meal_type=False):
    tag = template_img.copy()
    draw = _ImageDraw.Draw(tag)
    dpi_x, dpi_y = template_img.info.get("dpi", (300, 300))
    margin_x = _mm_to_px(BORDER_MARGIN_MM, dpi_x)
    margin_y = _mm_to_px(BORDER_MARGIN_MM, dpi_y)
    area_w, area_h = tag.width - 2 * margin_x, tag.height - 2 * margin_y

    wrapped, line_spacing_px = [], int(LINE_SPACING_FACTOR * max(MEAL_FONT_SIZES.values()))
    labels = ["dish", "meal", "client", "remarks"]

    for i, (t, label) in enumerate(zip(texts, labels)):
        if skip_all:
            if label != "dish":
                continue
        else:
            if label == "client" and not texts[3].strip():
                continue
            if label == "meal" and skip_meal_type:
                continue
            if label == "meal":
                t = _clean_meal_type(t)
        font = _ImageFont.truetype(font_path, MEAL_FONT_SIZES[label])
        wrapped.append((_wrap_line(draw, t, font, area_w), font))

    total_h = sum(len(lines) * (font.getbbox("Ag")[3] - font.getbbox("Ag")[1]) +
                  (len(lines) - 1) * line_spacing_px
                  for lines, font in wrapped)

    y = margin_y + (area_h - total_h) // 2
    for lines, font in wrapped:
        line_h = font.getbbox("Ag")[3] - font.getbbox("Ag")[1]
        for line in lines:
            x = margin_x + (area_w - draw.textlength(line, font=font)) // 2
            draw.text((x, y), line, fill=TAG_TEXT_COLOR, font=font)
            y += line_h + line_spacing_px
    return tag

def _draw_carrybag_tag(template_img, client, dishes, meal_type, remarks, slot, font_path):

    tag = template_img.copy()
    draw = _ImageDraw.Draw(tag)

    dpi_x, dpi_y = template_img.info.get("dpi", (300,300))

    margin_x = _mm_to_px(BORDER_MARGIN_MM, dpi_x)
    margin_y = _mm_to_px(BORDER_MARGIN_MM, dpi_y)

    area_w = tag.width - 2*margin_x
    area_h = tag.height - 2*margin_y
    scale = 1.0


    while True:

        wrapped = []

        if client:
            client_size = CARRYBAG_FONT_SIZES["client"]

            # shrink client name if many lines
            estimated_lines = len(dishes) + 3
            if estimated_lines > 6:
                client_size = int(client_size * 0.85)
            
            fnt_client = _ImageFont.truetype(font_path, int(client_size * scale))
            wrapped.append((_wrap_line(draw, client, fnt_client, area_w), fnt_client))

        for dish in dishes:
            fnt_dish = _ImageFont.truetype(font_path, int(CARRYBAG_FONT_SIZES["dish"] * scale))
            wrapped.append((_wrap_line(draw, dish, fnt_dish, area_w), fnt_dish))

        if meal_type:
            fnt_meal = _ImageFont.truetype(font_path, int(CARRYBAG_FONT_SIZES["meal"] * scale))
            wrapped.append((_wrap_line(draw, _clean_meal_type(meal_type), fnt_meal, area_w), fnt_meal))

        if remarks:
            fnt_rem = _ImageFont.truetype(font_path, int(CARRYBAG_FONT_SIZES["remarks"] * scale))
            wrapped.append((_wrap_line(draw, remarks, fnt_rem, area_w), fnt_rem))

        if slot and slot.strip().lower() != "afternoon":
            fnt_slot = _ImageFont.truetype(font_path, int(CARRYBAG_FONT_SIZES["remarks"] * scale))
            wrapped.append((_wrap_line(draw, slot, fnt_slot, area_w), fnt_slot))

        total_h = 0
        line_spacing_px = int(LINE_SPACING_FACTOR * dish_size)

        for lines, fnt in wrapped:
            line_h = fnt.getbbox("Ag")[3] - fnt.getbbox("Ag")[1]
            total_h += len(lines)*line_h + (len(lines)-1)*line_spacing_px

        if total_h <= area_h:
            break

        scale -= 0.05
        if scale < 0.6:
            break

    y = margin_y + (area_h-total_h)//2

    for lines, fnt in wrapped:

        line_h = fnt.getbbox("Ag")[3] - fnt.getbbox("Ag")[1]

        for line in lines:

            x = margin_x + (area_w-draw.textlength(line, font=fnt))//2
            draw.text((x,y), line, fill=TAG_TEXT_COLOR, font=fnt)

            y += line_h + line_spacing_px

    return tag

def run_tag_generator_auto_bytes(authed: AuthorizedSession, spreadsheet_id: str, sheet_name_for_tags: str,
                                 start_row: int, end_row: int) -> Tuple[bytes, int, int]:
    try:
        def _norm(x) -> str:
            return "" if x is None else str(x).strip()

        dish_map = _load_dish_map(DISHES_FILE)

        data = get_values_block(authed, spreadsheet_id, sheet_name_for_tags, "D", "M", start_row, end_row)

        font_path = _choose_font_path()

        meal_tags_data = []
        carry_groups: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

        buckets: Dict[int, List[Dict[str, Any]]] = {
            COL_BREAKFAST: [], COL_OPT1: [], COL_OPT2: [],
            COL_SNACK: [], COL_JUICE1: [], COL_JUICE2: []
        }

        TYPE_PRIORITY = ["chicken", "veg", "egg", "seafood"]

        row_counter = start_row - 1
        for row in data:
            row_counter += 1
            row_full = [str(x).strip() if x is not None else "" for x in (row + [""] * 20)]
            client_raw = row_full[COL_CLIENT] if len(row_full) > COL_CLIENT else ""
            client = re.sub(r"\s+", " ", client_raw).strip()
            if not client:
                continue

            meal_type = row_full[COL_TYPE] if len(row_full) > COL_TYPE else ""
            remarks = row_full[COL_REMARKS] if len(row_full) > COL_REMARKS else ""
            slot = row_full[COL_SLOT] if len(row_full) > COL_SLOT else ""

            row_dishes_codes = []
            for col_idx in [COL_BREAKFAST, COL_OPT1, COL_OPT2, COL_SNACK, COL_JUICE1, COL_JUICE2]:
                if col_idx < len(row_full) and str(row_full[col_idx]).strip():
                    code = str(row_full[col_idx]).strip()
                    dish_name = dish_map.get(code.lower(), code)

                    if col_idx in [COL_SNACK, COL_JUICE1, COL_JUICE2]:
                        skip_all = True
                        skip_type = True
                    else:
                        skip_all = False
                        skip_type = False

                    buckets[col_idx].append({
                        "code": code,
                        "dish": dish_name,
                        "orig_index": row_counter,
                        "client": client,
                        "type": meal_type,
                        "remarks": remarks,
                        "skip_all": skip_all,
                        "skip_type": skip_type
                    })

                    row_dishes_codes.append((code, col_idx))

            if row_dishes_codes:
                key = (client.strip().lower(), (meal_type or "").strip().lower(), (slot or "").strip().lower())
                if key not in carry_groups:
                    carry_groups[key] = {
                        "client": client,
                        "type": meal_type,
                        "remarks": remarks,
                        "counts": {},
                        "order": [],
                        "cols": {}
                    }

                grp = carry_groups[key]
                for code, col_idx in row_dishes_codes:
                    if code not in grp["counts"]:
                        grp["counts"][code] = 0
                        grp["order"].append(code)
                        grp["cols"][code] = col_idx
                    grp["counts"][code] += 1

                if not grp["remarks"] and remarks:
                    grp["remarks"] = remarks

        def type_rank_for_code(code: str) -> int:
            cl = code.lower()
            for idx, t in enumerate(TYPE_PRIORITY):
                if cl.startswith(t) or f" {t}" in cl or cl.startswith(t + "-") or t in cl:
                    return idx
            return len(TYPE_PRIORITY) + 1

        priority_cols = [COL_BREAKFAST, COL_OPT1, COL_OPT2, COL_SNACK, COL_JUICE1, COL_JUICE2]
        meal_tags_data = []
        for pc in priority_cols:
            bucket = buckets.get(pc, [])
            bucket_sorted = sorted(bucket, key=lambda itm: (type_rank_for_code(itm["code"]), itm["orig_index"]))
            for itm in bucket_sorted:
                meal_tags_data.append({
                    "dish": itm["dish"],
                    "client": itm["client"],
                    "type": itm["type"],
                    "remarks": itm["remarks"],
                    "skip_all": itm["skip_all"],
                    "skip_type": itm["skip_type"]
                })

        carrybag_tags_data = []
        for (_k_client, _k_type, _k_slot), group in carry_groups.items():
            counts = group["counts"]
            def carry_sort_key(code):
                col_idx = group["cols"].get(code, 999)
                try:
                    col_priority = priority_cols.index(col_idx)
                except ValueError:
                    col_priority = 999
                typ_pri = type_rank_for_code(code)
                orig_pos = group["order"].index(code) if code in group["order"] else 9999
                return (col_priority, typ_pri, orig_pos)
            ordered_codes = sorted(list(group["order"]), key=carry_sort_key)
            dishes_display = [
                (f"{counts[code]}x {code}" if counts[code] > 1 else code)
                for code in ordered_codes
            ]
            carrybag_tags_data.append({
                "client": group["client"],
                "dishes": dishes_display,
                "type": group["type"],
                "remarks": group["remarks"],
                "slot": _k_slot
            })

        meal_count = len(meal_tags_data)
        carrybag_count = len(carrybag_tags_data)

        meal_template = _Image.open(MEAL_TEMPLATE_FILE).convert("RGB")
        carrybag_template = _Image.open(CARRYBAG_TEMPLATE_FILE).convert("RGB")

        all_tags_imgs = []
        for tag_data in meal_tags_data:
            all_tags_imgs.append(_draw_meal_tag(
                meal_template,
                [tag_data["dish"], tag_data["type"], tag_data["client"], tag_data["remarks"]],
                _choose_font_path(),
                skip_all=tag_data["skip_all"], skip_meal_type=tag_data["skip_type"]
            ))
        for tag_data in carrybag_tags_data:
            all_tags_imgs.append(_draw_carrybag_tag(
                carrybag_template,
                tag_data["client"], tag_data["dishes"], tag_data["type"], tag_data["remarks"],
                tag_data["slot"],
                _choose_font_path()
            ))

        # Render PDF entirely in-memory
        buf = BytesIO()
        PAGE_W, PAGE_H = _A4
        c = _canvas.Canvas(buf, pagesize=_A4)

        def layout_info(template_img):
            dpi_x, dpi_y = template_img.info.get("dpi", (300, 300))
            tag_w_pt = template_img.width * 72.0 / dpi_x
            tag_h_pt = template_img.height * 72.0 / dpi_y
            cols = max(1, int(round(PAGE_W / tag_w_pt)))
            rows_per_page = max(1, int(round(PAGE_H / tag_h_pt)))
            per_page = cols * rows_per_page
            return tag_w_pt, tag_h_pt, cols, rows_per_page, per_page

        tag_w_pt, tag_h_pt, cols, rows_per_page, per_page = layout_info(meal_template)

        total_imgs = len(all_tags_imgs)
        for idx, img in enumerate(all_tags_imgs):
            if idx and idx % per_page == 0:
                c.showPage()
            col = (idx % per_page) % cols
            row = (idx % per_page) // cols
            x = col * tag_w_pt
            y = PAGE_H - (row + 1) * tag_h_pt
            c.drawImage(_ImageReader(img), x, y, width=tag_w_pt, height=tag_h_pt)

        # Fill remaining slots on last page with blank meal template
        if total_imgs:
            remaining = (per_page - (total_imgs % per_page)) % per_page
            for i in range(remaining):
                idx = total_imgs + i
                if idx and idx % per_page == 0:
                    c.showPage()
                col = (idx % per_page) % cols
                row = (idx % per_page) // cols
                x = col * tag_w_pt
                y = PAGE_H - (row + 1) * tag_h_pt
                c.drawImage(_ImageReader(meal_template), x, y, width=tag_w_pt, height=tag_h_pt)

        c.save()
        pdf_bytes = buf.getvalue()
        buf.close()
        return pdf_bytes, meal_count, carrybag_count

    except Exception as e:
        raise SystemExit(f"Tag generation error: {e}")

# ============================================================
# SILENT CSV GENERATOR (from sheet "Menu" – case-insensitive)
# ============================================================

def _norm(x) -> str:
    return "" if x is None else str(x).strip()

def _collect_code_columns_from_row2(values: List[List[str]]) -> List[int]:
    if len(values) < 2:
        return []
    header = values[1]
    cols = []
    max_idx = max(len(header), 3)
    for c in range(3, max_idx):
        label = _norm(header[c] if c < len(header) else "")
        if label:
            cols.append(c)
    return cols

def _rows_matching_date_in_colC(values: List[List[str]], normalized_date: str) -> List[int]:
    matches = []
    for r, row in enumerate(values):
        if len(row) <= 2:
            continue
        if as_sheet_date_text(row[2]) == normalized_date:
            matches.append(r)
    return matches

def _export_csv(path: str, rows: List[List[str]]) -> None:
    import csv
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)

def generate_dishes_csv_for_date(authed: AuthorizedSession, spreadsheet_id: str, normalized_date: str, output_csv: str = "dishes.csv") -> None:
    menu_title = resolve_sheet_title_ci(authed, spreadsheet_id, "Menu")
    values = get_values_safe(authed, spreadsheet_id, f"{menu_title}!A1:ZZ10000")
    if not values:
        _export_csv(output_csv, [["Code", "DishName"]]); return
    code_cols = _collect_code_columns_from_row2(values)
    if not code_cols:
        _export_csv(output_csv, [["Code", "DishName"]]); return
    header_row = values[1] if len(values) > 1 else []
    matches = _rows_matching_date_in_colC(values, normalized_date)
    out = [["Code", "DishName"]]
    for r in matches:
        if r <= 1:
            continue
        row = values[r]
        for c in code_cols:
            code = _norm(header_row[c] if c < len(header_row) else "")
            dish = _norm(row[c] if c < len(row) else "")
            if code and dish:
                out.append([code, dish])
    _export_csv(output_csv, out)

# ============================================================
# MAIN
# ============================================================

def _orient_char_to_bool_strict(ch: str) -> bool:
    c = (ch or "").strip().lower()
    if c == 'p':
        return True
    if c == 'l':
        return False
    raise ValueError("Orientation must be 'p' (portrait) or 'l' (landscape).")

def _parse_month_input(mstr: str, default_month_num: int) -> Tuple[int, str]:
    """
    Accepts: "", "November", "nov", "11"
    Returns: (month_number, MonthName). If blank -> (default_month_num, DefaultMonthName).
    """
    mstr = (mstr or "").strip()
    if not mstr:
        return default_month_num, calendar.month_name[default_month_num]

    # numeric?
    try:
        mnum = int(mstr)
        if 1 <= mnum <= 12:
            return mnum, calendar.month_name[mnum]
    except ValueError:
        pass

    # names/abbr
    names = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}
    abbrs = {name.lower(): i for i, name in enumerate(calendar.month_abbr) if name}
    key = mstr.lower()
    if key in names:
        i = names[key]
        return i, calendar.month_name[i]
    if key in abbrs:
        i = abbrs[key]
        return i, calendar.month_name[i]

    raise ValueError(f"Unrecognized month: {mstr}")

def main():
    creds = get_creds()
    authed = AuthorizedSession(creds)
    spreadsheet_id = get_spreadsheet_id_from_url(SHEET_URL)

    # Resolve orientations (validation only — no prompt shown)
    try:
        p1_portrait = _orient_char_to_bool_strict(PAGE1_ORIENT)
        p2_portrait = _orient_char_to_bool_strict(PAGE2_ORIENT)
        p3_portrait = _orient_char_to_bool_strict(PAGE3_ORIENT)
        p4_portrait = _orient_char_to_bool_strict(PAGE4_ORIENT)
    except ValueError as e:
        raise SystemExit(f"Invalid PAGE*_ORIENT value: {e}")

    _bar(0)

    # ---------- Month first (with memory), then Day ----------
    last = load_last_used()
    default_month_num = None
    if isinstance(last, dict) and last.get("month"):
        try:
            default_month_num, _ = _parse_month_input(last["month"], date.today().month)
        except Exception:
            default_month_num = None

    if default_month_num is None and isinstance(last, dict) and last.get("date"):
        try:
            tmp_norm = normalize_input_date(last["date"])
            default_month_num = datetime.strptime(tmp_norm, "%d-%b-%y").month
        except Exception:
            default_month_num = None

    if default_month_num is None:
        default_month_num = date.today().month

    default_month_name = calendar.month_name[default_month_num]

    print("\nEnter month (name/abbr/number).")
    print(f"Press Enter to use: {default_month_name}")
    month_input = input("Month: ").strip()
    try:
        month_num, month_name = _parse_month_input(month_input, default_month_num)
    except ValueError as e:
        _bar(0)
        raise SystemExit(str(e))

    save_last_used(month_name=month_name)

    # Choose year: try to reuse last date's year, else current year
    default_year = date.today().year
    if isinstance(last, dict) and last.get("date"):
        try:
            tmp_norm = normalize_input_date(last["date"])
            default_year = datetime.strptime(tmp_norm, "%d-%b-%y").year
        except Exception:
            pass

    day_str = input("Service day (1-31): ").strip()
    if not re.fullmatch(r"\d{1,2}", day_str):
        raise SystemExit("Please enter a valid day number (1-31).")
    day_num = int(day_str)
    if not (1 <= day_num <= 31):
        raise SystemExit("Day must be in 1..31.")

    try:
        service_dt = datetime(default_year, month_num, day_num)
    except ValueError:
        service_dt = datetime(default_year, month_num, 1)

    normalized_date = service_dt.strftime("%d-%b-%y")
    try:
        normalized_date = normalize_input_date(f"{day_num}-{month_name[:3]}-{service_dt.year}")
    except Exception:
        pass

    save_last_used(date_in=normalized_date)

    final_output_pdf = f"{normalized_date} list.pdf"

    _bar(5)

    # 0) SILENT CSV from "Menu"
    generate_dishes_csv_for_date(authed, spreadsheet_id, normalized_date, output_csv="dishes.csv")
    _bar(25)

    # 1) Pages 1 & 2 from Dailylist (accept full or abbr month in tab)
    dailylist_title = resolve_dailylist_for_month(authed, spreadsheet_id, month_name)

    start_row, next_row = find_start_and_next_rows(authed, spreadsheet_id, dailylist_title, normalized_date)
    merges_12 = get_sheet_merges(authed, spreadsheet_id, dailylist_title)
    _bar(40)

    end_cm = find_end_by_two_empty_rows(
        authed, spreadsheet_id, dailylist_title, "C", "M", start_row,
        max_row=9999, empty_run=2, merges=merges_12
    )
    if next_row is not None:
        end_pz = max(start_row, next_row - 3)
    else:
        end_pz = find_last_used_row_pz_till_sheet_end(
            authed,
            spreadsheet_id,
            dailylist_title,
            start_row
    )

    _bar(50)

    range1 = f"C{start_row}:M{end_cm}"  # Page 1 (Client list)
    range2 = f"P{start_row}:Z{end_pz}"  # Page 2 (Meal list)

    gid_12 = get_gid_for_sheet(dailylist_title, spreadsheet_id, authed)

    # --- Page 1 export with manual one-page toggle ---
    r1, c1, r2, c2 = a1_to_rc(range1)
    p1_bytes = export_range_pdf_bytes(
        spreadsheet_id, gid_12, r1, c1, r2, c2,
        authed, portrait=p1_portrait,
        fit_page=PAGE1_FORCE_ONE_PAGE
    )
    _bar(65)

    # --- Page 2 export: try fit-to-width, fallback to fit-to-page if it spills ---
    r1, c1, r2, c2 = a1_to_rc(range2)
    p2_bytes = export_range_pdf_bytes(
        spreadsheet_id, gid_12, r1, c1, r2, c2,
        authed, portrait=p2_portrait, fit_page=False
    )
    try:
        reader = PdfReader(BytesIO(p2_bytes))
        page_count = len(reader.pages)
    except Exception:
        page_count = 2

    if page_count > 1:
        p2_bytes = export_range_pdf_bytes(
            spreadsheet_id, gid_12, r1, c1, r2, c2,
            authed, portrait=p2_portrait, fit_page=False
        )
    _bar(75)

    # 2) Page 3 from "MiseList"
    page3_sheet_title, page3_a1 = auto_find_page3_range_on_miselist(authed, spreadsheet_id, normalized_date)
    gid_3  = get_gid_for_sheet(page3_sheet_title, spreadsheet_id, authed)
    r1, c1, r2, c2 = a1_to_rc(page3_a1)
    p3_bytes = export_range_pdf_bytes(
        spreadsheet_id, gid_3, r1, c1, r2, c2,
        authed, portrait=p3_portrait, fit_page=True
    )
    _bar(85)

    # 3) Delivery page (optional)
    p_delivery_bytes = None
    if DELIVERY_ENABLED:
        try:
            delivery_found = auto_find_delivery_block_range(authed, spreadsheet_id, normalized_date)
            if delivery_found:
                delivery_sheet_title, delivery_a1 = delivery_found
                gid_delivery = get_gid_for_sheet(delivery_sheet_title, spreadsheet_id, authed)
                r1, c1, r2, c2 = a1_to_rc(delivery_a1)
                p_delivery_bytes = export_range_pdf_bytes(
                    spreadsheet_id, gid_delivery, r1, c1, r2, c2,
                    authed, portrait=p4_portrait, fit_page=True
                )
            else:
                from reportlab.pdfgen import canvas as __canvas
                from reportlab.lib.pagesizes import A4 as __A4
                tmp_buf = BytesIO()
                ctmp = __canvas.Canvas(tmp_buf, pagesize=__A4)
                ctmp.showPage()
                ctmp.save()
                p_delivery_bytes = tmp_buf.getvalue()
                tmp_buf.close()
                print(f"\nNotice: No Delivery block found for {normalized_date}; skipping Delivery content.")
        except Exception as e:
            from reportlab.pdfgen import canvas as __canvas
            from reportlab.lib.pagesizes import A4 as __A4
            tmp_buf = BytesIO()
            ctmp = __canvas.Canvas(tmp_buf, pagesize=__A4)
            ctmp.showPage()
            ctmp.save()
            p_delivery_bytes = tmp_buf.getvalue()
            tmp_buf.close()
            print(f"\nWarning: Delivery export failed: {e}")
    _bar(90)

    # Tags from Page 1 rows
    tag_start_row = start_row + 2
    tag_end_row   = end_cm

    if tag_start_row <= tag_end_row:
        tags_bytes, meal_count, carrybag_count = run_tag_generator_auto_bytes(
            authed, spreadsheet_id, dailylist_title, tag_start_row, tag_end_row
        )
    else:
        from reportlab.pdfgen import canvas as __canvas
        from reportlab.lib.pagesizes import A4 as __A4
        tmp_buf = BytesIO()
        c = __canvas.Canvas(tmp_buf, pagesize=__A4)
        c.showPage()
        c.save()
        tags_bytes = tmp_buf.getvalue()
        tmp_buf.close()
        meal_count = 0
        carrybag_count = 0
    _bar(95)

    # Merge to final: TAGS FIRST, then Page1, Page2, Page3, Delivery (if present)
    pdfs_to_merge = [tags_bytes, p1_bytes, p2_bytes, p3_bytes]
    if DELIVERY_ENABLED and p_delivery_bytes is not None:
        pdfs_to_merge.append(p_delivery_bytes)

    merge_pdfs_bytes_to_file(pdfs_to_merge, final_output_pdf)
    _bar(100)

    total_tags = meal_count + carrybag_count
    print(f"\nMeal tags: {meal_count} | Carrybag tags: {carrybag_count} | Total: {total_tags}")

    try:
        input("\nPress Enter to exit...")
    except Exception:
        pass

if __name__ == "__main__":
    main()
