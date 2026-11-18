"""Confirmed company names (Vietnamese -> international name), shared by all users.

Stored locally in company_names.csv next to the app, and on Streamlit Cloud in a "CompanyNames" worksheet of the
logging spreadsheet (files written at runtime on Streamlit Cloud are lost on restart). Rows are append-only, which
keeps an audit trail; the latest row for a name wins, and a row with an empty international name removes it.
"""
import csv
import os
import unicodedata
from datetime import datetime, timedelta, timezone

import streamlit as st

import usage_logger


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_FILE = os.path.join(BASE_DIR, "company_names.csv")
SHEET_TITLE = "CompanyNames"
HEADERS = ["Vietnamese", "International", "Source", "Saved by", "Saved at"]


def normalize_name(name):
    return " ".join(unicodedata.normalize("NFC", str(name)).split())


def _worksheet(create=False):
    import gspread

    client, url_or_error = usage_logger.get_gsheet_client()
    if not client:
        raise RuntimeError(f"Không kết nối được Google Sheet: {url_or_error}")
    spreadsheet = client.open_by_url(url_or_error)
    try:
        return spreadsheet.worksheet(SHEET_TITLE)
    except gspread.WorksheetNotFound:
        if not create:
            return None
        worksheet = spreadsheet.add_worksheet(title=SHEET_TITLE, rows=1000, cols=len(HEADERS))
        worksheet.append_row(HEADERS)
        return worksheet


def _read_rows():
    if usage_logger.is_local_env():
        if not os.path.exists(LOCAL_FILE):
            return []
        with open(LOCAL_FILE, encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    worksheet = _worksheet()
    return worksheet.get_all_records() if worksheet is not None else []


@st.cache_data(ttl=300, show_spinner=False)
def _cached_rows():
    return _read_rows()


def clear_cache():
    _cached_rows.clear()


def load_company_names(use_cache=True):
    """{normalized Vietnamese name: {"vietnamese", "international", "source", "saved_by", "saved_at"}}."""
    names = {}
    for row in (_cached_rows() if use_cache else _read_rows()):
        vietnamese = normalize_name(row.get("Vietnamese", ""))
        if not vietnamese:
            continue
        international = str(row.get("International", "")).strip()
        if not international:
            names.pop(vietnamese, None)
            continue
        names[vietnamese] = {"vietnamese": vietnamese, "international": international, "source": str(row.get("Source", "")),
                             "saved_by": str(row.get("Saved by", "")), "saved_at": str(row.get("Saved at", ""))}
    return names


def company_name_map(use_cache=True):
    """{Vietnamese name: international name}, ready to merge into a translation map."""
    return {entry["vietnamese"]: entry["international"] for entry in load_company_names(use_cache).values()}


def save_company_names(entries, user):
    """Append confirmed names. entries: [{"vietnamese", "international", "source"}]; returns the number saved."""
    timestamp = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for entry in entries:
        vietnamese = normalize_name(entry.get("vietnamese", ""))
        international = " ".join(str(entry.get("international", "") or "").split())
        if vietnamese and international:
            rows.append([vietnamese, international, str(entry.get("source", "") or ""), str(user), timestamp])
    if not rows:
        return 0
    if usage_logger.is_local_env():
        new_file = not os.path.exists(LOCAL_FILE)
        with open(LOCAL_FILE, "a", encoding="utf-8-sig" if new_file else "utf-8", newline="") as handle:
            writer = csv.writer(handle)
            if new_file:
                writer.writerow(HEADERS)
            writer.writerows(rows)
    else:
        _worksheet(create=True).append_rows(rows, value_input_option="RAW")
    clear_cache()
    return len(rows)


def dictionary_rows(use_cache=True):
    """Confirmed names in the Dictionary_v3 layout (Vietnamese, E, Hs, Ht), for merging into the dictionary."""
    return [{"Vietnamese": vietnamese, "E": international, "Hs": international, "Ht": international}
            for vietnamese, international in sorted(company_name_map(use_cache).items())]
