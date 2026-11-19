import csv
import os
import time
import random
import re
from datetime import datetime, timedelta, timezone
import streamlit as st

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def is_local_env():
    """Detects if the app is running locally (Windows) or on Streamlit Cloud (Linux)."""
    # Windows is always local host for this user.
    if os.name == 'nt':
        return True
    # Streamlit Cloud runs on Linux and sets specific env vars.
    if os.environ.get("STREAMLIT_SHARING_MODE") is not None:
        return False
    # Fallback to the original method
    return hasattr(os, 'startfile')

# --- Google Sheets Logic ---
def normalize_private_key(key):
    """Rebuild a PEM private key pasted in a slightly wrong shape: literal or doubled \\n, Windows line ends,
    indented lines, stray quotes, or the whole key on one line with spaces instead of line breaks."""
    text = key.replace("\\\\n", "\n").replace("\\n", "\n").replace("\r", "")
    match = re.search(r"-----BEGIN ([A-Z ]+)-----(.*?)-----END \1-----", text, re.S)
    if not match:
        return text
    body = re.sub(r"[^A-Za-z0-9+/=]", "", match.group(2))
    lines = [body[i:i + 64] for i in range(0, len(body), 64)]
    return "\n".join([f"-----BEGIN {match.group(1)}-----", *lines, f"-----END {match.group(1)}-----"]) + "\n"


def get_gsheet_client():
    """Initializes gspread client using Streamlit secrets."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        # Support both [gsheets] and [connections.gsheets] structures
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            creds_dict = dict(st.secrets["connections"]["gsheets"])
        elif "gsheets" in st.secrets:
            creds_dict = dict(st.secrets["gsheets"])
        else:
            return None, "Secrets not found in st.secrets"
        
        # Support both 'spreadsheet' (standard) and 'spreadsheet_url' (custom)
        spreadsheet_url = creds_dict.pop("spreadsheet", None) or creds_dict.pop("spreadsheet_url", None)
        
        if not spreadsheet_url:
            return None, "Spreadsheet URL (spreadsheet) not found in secrets"
            
        if "private_key" in creds_dict and isinstance(creds_dict["private_key"], str):
            creds_dict["private_key"] = normalize_private_key(creds_dict["private_key"])

        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        return client, spreadsheet_url
    except Exception as e:
        return None, f"Client init error: {str(e)}"

@st.cache_resource(show_spinner=False)
def _cached_spreadsheet():
    client, url_or_error = get_gsheet_client()
    if not client:
        raise RuntimeError(url_or_error)
    return client.open_by_url(url_or_error)


def open_spreadsheet():
    """The logging spreadsheet, opened once and reused (a failed attempt is not cached, so it is retried)."""
    try:
        return _cached_spreadsheet()
    except Exception as e:
        raise RuntimeError(f"Không kết nối được Google Sheet: {e}") from e


def _timestamp():
    return datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M:%S")


def log_to_gsheet(user, event_type, details):
    """Appends a log entry to Google Sheets, retrying on rate limits. Returns True when written."""
    row = [_timestamp(), user, event_type, details]
    for attempt in range(3):
        try:
            open_spreadsheet().get_worksheet(0).append_row(row)
            return True
        except Exception as e:
            if attempt == 2 or "Không kết nối được" in str(e):
                st.session_state["gsheet_error"] = str(e)
                print(f"GSheet logging failed: {e}")
                return False
            # Exponential backoff with jitter to reduce congestion
            time.sleep(1 + attempt + random.random())
    return False

def get_gsheet_logs():
    """Fetches all logs from Google Sheets (newest first); raises when the sheet cannot be read."""
    return open_spreadsheet().get_worksheet(0).get_all_records()[::-1]


def _log_to_csv(user, event_type, details):
    file_exists = os.path.isfile(LOG_FILE)
    try:
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "User", "Event Type", "Details"])
            writer.writerow([_timestamp(), user, event_type, details])
    except Exception as e:
        print(f"Error logging event: {e}")


def _csv_logs():
    if not os.path.exists(LOG_FILE):
        return []
    try:
        with open(LOG_FILE, mode='r', encoding='utf-8') as f:
            return list(csv.DictReader(f))[::-1]
    except Exception:
        return []

# --- Public API ---

def log_event(user, event_type, details):
    """Logs to Google Sheets on every machine; falls back to usage_log.csv when the sheet is unreachable."""
    if not log_to_gsheet(user, event_type, details):
        _log_to_csv(user, event_type, details)

def get_logs():
    """Logs from Google Sheets; usage_log.csv only when the sheet cannot be read."""
    try:
        logs = get_gsheet_logs()
        st.session_state.pop("gsheet_read_error", None)
        return logs
    except Exception as e:
        st.session_state["gsheet_read_error"] = str(e)
        return _csv_logs()
