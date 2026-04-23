import csv
import os
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

# --- Google Sheets Logic (Cloud) ---
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
            
        # Fix private key formatting (replace literal \n with actual newlines if necessary)
        if "private_key" in creds_dict and isinstance(creds_dict["private_key"], str):
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        return client, spreadsheet_url
    except Exception as e:
        return None, f"Client init error: {str(e)}"

def log_to_gsheet(user, event_type, details):
    """Appends a log entry to Google Sheets."""
    client, sheet_url_or_err = get_gsheet_client()
    if not client:
        print(f"GSheet client failed: {sheet_url_or_err}")
        return
    
    try:
        spreadsheet = client.open_by_url(sheet_url_or_err)
        worksheet = spreadsheet.get_worksheet(0)
        
        # Check if sheet is empty and add headers
        if not worksheet.get_all_values():
            worksheet.append_row(["Timestamp", "User", "Event Type", "Details"])

        vn_tz = timezone(timedelta(hours=7))
        timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
        
        worksheet.append_row([timestamp, user, event_type, details])
    except Exception as e:
        # Store error in session state for UI display if possible
        st.session_state["gsheet_error"] = str(e)
        print(f"GSheet logging failed: {e}")

def get_gsheet_logs():
    """Fetches all logs from Google Sheets."""
    client, sheet_url = get_gsheet_client()
    if not client or not sheet_url:
        return []
    
    try:
        spreadsheet = client.open_by_url(sheet_url)
        worksheet = spreadsheet.get_worksheet(0)
        data = worksheet.get_all_records()
        return data[::-1] # Reverse for newest first
    except Exception:
        return []

# --- Public API ---

def log_event(user, event_type, details):
    """Logs an event to the appropriate destination based on environment."""
    if is_local_env():
        # Local logging to CSV
        vn_tz = timezone(timedelta(hours=7))
        timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
        file_exists = os.path.isfile(LOG_FILE)
        
        try:
            with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["Timestamp", "User", "Event Type", "Details"])
                writer.writerow([timestamp, user, event_type, details])
        except Exception as e:
            print(f"Error logging event: {e}")
    else:
        # Cloud logging to Google Sheets
        log_to_gsheet(user, event_type, details)

def get_logs():
    """Reads logs from the appropriate destination based on environment."""
    if is_local_env():
        if not os.path.exists(LOG_FILE):
            return []
        try:
            logs = []
            with open(LOG_FILE, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    logs.append(row)
            return logs[::-1]
        except Exception:
            return []
    else:
        # Fetch from Google Sheets
        return get_gsheet_logs()
