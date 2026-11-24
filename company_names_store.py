"""Confirmed company names (Vietnamese -> international name), shared by all users.

Stored in a "CompanyNames" worksheet of the logging Google Sheet, so names can be tracked and reviewed in one place.
When the Sheet is not configured or cannot be reached on a local (Windows) run, names fall back to company_names.csv
next to the app; upload_local_to_sheet() moves them to the Sheet later.

Rows are append-only, which keeps an audit trail. For each name the row with the latest "Saved at" wins (so sorting or
filtering the Sheet does not change the result), and a row with an empty international name removes the name.
Reviewers fill Status / Reviewed by / Note directly in the Sheet; names are used as soon as they are saved (Status is
for tracking only), and a corrected international name takes effect within the read cache time (5 minutes).
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
HEADERS = ["Vietnamese", "International", "Source", "Saved by", "Saved at", "Status", "Reviewed by", "Note", "Tax code"]
STATUS_PENDING = "Chờ duyệt"
STATUS_APPROVED = "Đã duyệt"
STATUS_REJECTED = "Không dùng"
STATUSES = (STATUS_PENDING, STATUS_APPROVED, STATUS_REJECTED)
FIELDS = {"vietnamese": "Vietnamese", "international": "International", "source": "Source", "saved_by": "Saved by",
          "saved_at": "Saved at", "status": "Status", "reviewed_by": "Reviewed by", "note": "Note", "tax_code": "Tax code"}


def normalize_name(name):
    return " ".join(unicodedata.normalize("NFC", str(name)).split())


# ---------------------------------------------------------------- where names are stored

def _spreadsheet():
    client, url_or_error = usage_logger.get_gsheet_client()
    if not client:
        raise RuntimeError(f"Không kết nối được Google Sheet: {url_or_error}")
    return client.open_by_url(url_or_error)


def _backend():
    """("sheet", spreadsheet) when the Google Sheet is reachable; ("csv", reason) as a local fallback; raises on the Cloud."""
    try:
        return "sheet", _spreadsheet()
    except Exception as error:
        if usage_logger.is_local_env():
            return "csv", str(error)
        raise RuntimeError(f"Không kết nối được Google Sheet: {error}") from error


def _format_sheet(spreadsheet, worksheet):
    """Bold, frozen header and a Status dropdown. Cosmetic: a failure here never blocks saving."""
    try:
        worksheet.freeze(rows=1)
        worksheet.format("A1:I1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.87, "green": 0.91, "blue": 0.95}})
        status_column = HEADERS.index("Status")
        spreadsheet.batch_update({"requests": [{"setDataValidation": {
            "range": {"sheetId": worksheet.id, "startRowIndex": 1, "startColumnIndex": status_column, "endColumnIndex": status_column + 1},
            "rule": {"condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": value} for value in STATUSES]},
                     "showCustomUi": True, "strict": False}}}]})
    except Exception as error:
        print(f"CompanyNames formatting skipped: {error}")


def _worksheet(spreadsheet, create=False):
    """The CompanyNames worksheet with the current header. An older 5-column header is extended in place."""
    import gspread

    try:
        worksheet = spreadsheet.worksheet(SHEET_TITLE)
    except gspread.WorksheetNotFound:
        if not create:
            return None
        worksheet = spreadsheet.add_worksheet(title=SHEET_TITLE, rows=1000, cols=len(HEADERS))
        worksheet.update("A1", [HEADERS], value_input_option="RAW")
        _format_sheet(spreadsheet, worksheet)
        return worksheet
    if create:
        header = [str(value).strip() for value in worksheet.row_values(1)]
        if header != HEADERS:
            if header and header != HEADERS[:len(header)]:
                raise ValueError(f"Dòng tiêu đề của sheet {SHEET_TITLE} đã bị sửa. Cần đúng thứ tự cột: " + ", ".join(HEADERS))
            if worksheet.col_count < len(HEADERS):
                worksheet.add_cols(len(HEADERS) - worksheet.col_count)
            worksheet.update("A1", [HEADERS], value_input_option="RAW")
            _format_sheet(spreadsheet, worksheet)
    return worksheet


def _sheet_rows(spreadsheet):
    worksheet = _worksheet(spreadsheet)
    if worksheet is None:
        return []
    values = worksheet.get_all_values()
    if not values:
        return []
    header = [str(value).strip() for value in values[0]]
    return [{name: (row[index] if index < len(row) else "") for index, name in enumerate(header) if name} for row in values[1:]]


def _csv_rows():
    if not os.path.exists(LOCAL_FILE):
        return []
    with open(LOCAL_FILE, encoding="utf-8-sig", newline="") as handle:
        return [{key: value or "" for key, value in row.items() if key} for row in csv.DictReader(handle)]


def _write_csv(rows):
    """Rewrite the local file with the current header (older files with fewer columns are upgraded)."""
    with open(LOCAL_FILE, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows([{name: row.get(name, "") for name in HEADERS} for row in rows])


def _read_rows():
    kind, target = _backend()
    return _sheet_rows(target) if kind == "sheet" else _csv_rows()


def storage_status():
    """Where names are read and saved now, for the management tab:
    {"backend": "sheet" | "csv", "url": sheet link or "", "detail": fallback reason or "", "local_rows": rows waiting in the CSV}."""
    kind, target = _backend()
    status = {"backend": kind, "url": "", "detail": "", "local_rows": 0}
    if kind == "sheet":
        worksheet = _worksheet(target)
        base = getattr(target, "url", "") or ""
        status["url"] = f"{base}#gid={worksheet.id}" if base and worksheet is not None else base
        status["local_rows"] = len(_csv_rows())
    else:
        status["detail"] = target
    return status


@st.cache_data(ttl=300, show_spinner=False)
def _cached_rows():
    return _read_rows()


def clear_cache():
    _cached_rows.clear()


# ---------------------------------------------------------------- reading and saving

def _entry(row):
    entry = {field: str(row.get(column, "") or "").strip() for field, column in FIELDS.items()}
    entry["vietnamese"] = normalize_name(entry["vietnamese"])
    entry["international"] = " ".join(entry["international"].split())
    entry["status"] = entry["status"] or STATUS_PENDING
    return entry


def load_company_names(use_cache=True):
    """{normalized Vietnamese name: {"vietnamese", "international", "source", "saved_by", "saved_at", "status",
    "reviewed_by", "note", "tax_code"}}. The latest "Saved at" wins; an empty international name removes the name."""
    rows = [_entry(row) for row in (_cached_rows() if use_cache else _read_rows())]
    names = {}
    for entry in sorted((entry for entry in rows if entry["vietnamese"]), key=lambda entry: entry["saved_at"]):
        if entry["international"]:
            names[entry["vietnamese"]] = entry
        else:
            names.pop(entry["vietnamese"], None)
    return names


def company_name_map(use_cache=True):
    """{Vietnamese name: international name}, ready to merge into a translation map. Names marked "Không dùng" are left out."""
    return {entry["vietnamese"]: entry["international"] for entry in load_company_names(use_cache).values()
            if entry["status"] != STATUS_REJECTED}


def _timestamp():
    return datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M:%S")


def save_company_names(entries, user):
    """Append confirmed names as "Chờ duyệt". entries: [{"vietnamese", "international", "source"?, "note"?, "tax_code"?}].
    Returns the number saved."""
    timestamp = _timestamp()
    rows = []
    for entry in entries:
        vietnamese = normalize_name(entry.get("vietnamese", ""))
        international = " ".join(str(entry.get("international", "") or "").split())
        if vietnamese and international:
            rows.append({"Vietnamese": vietnamese, "International": international, "Source": str(entry.get("source", "") or ""),
                         "Saved by": str(user), "Saved at": timestamp, "Status": STATUS_PENDING, "Reviewed by": "",
                         "Note": str(entry.get("note", "") or ""), "Tax code": str(entry.get("tax_code", "") or "")})
    if not rows:
        return 0
    kind, target = _backend()
    if kind == "sheet":
        _worksheet(target, create=True).append_rows([[row[name] for name in HEADERS] for row in rows], value_input_option="RAW")
    else:
        _write_csv(_csv_rows() + rows)
    clear_cache()
    return len(rows)


def upload_local_to_sheet():
    """Move the names saved in company_names.csv (local fallback) to the Google Sheet, skipping rows already there.
    The CSV is then renamed to company_names_uploaded_<time>.csv as a backup. Returns the number of rows added."""
    kind, target = _backend()
    if kind != "sheet":
        raise RuntimeError(f"Chưa kết nối được Google Sheet: {target}")
    local = _csv_rows()
    if not local:
        return 0
    worksheet = _worksheet(target, create=True)

    def key(row):
        return normalize_name(row.get("Vietnamese", "")), " ".join(str(row.get("International", "")).split()), str(row.get("Saved at", "")).strip()

    existing = {key(row) for row in _sheet_rows(target)}
    new_rows = [row for row in local if key(row) not in existing and normalize_name(row.get("Vietnamese", ""))]
    if new_rows:
        worksheet.append_rows([[str(row.get(name, "") or (STATUS_PENDING if name == "Status" else "")) for name in HEADERS] for row in new_rows],
                              value_input_option="RAW")
    stamp = datetime.now(timezone(timedelta(hours=7))).strftime("%Y%m%d_%H%M%S")
    os.replace(LOCAL_FILE, os.path.join(os.path.dirname(LOCAL_FILE), f"company_names_uploaded_{stamp}.csv"))
    clear_cache()
    return len(new_rows)


def dictionary_rows(use_cache=True):
    """Confirmed names in the Dictionary_v3 layout (Vietnamese, E, Hs, Ht), for merging into the dictionary."""
    return [{"Vietnamese": vietnamese, "E": international, "Hs": international, "Ht": international}
            for vietnamese, international in sorted(company_name_map(use_cache).items())]
