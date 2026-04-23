import csv
import os
from datetime import datetime, timedelta, timezone
import streamlit as st

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def is_local_env():
    """Detects if the app is running locally (Windows) or on Streamlit Cloud (Linux)."""
    # Local host (Windows) has os.startfile, Streamlit Cloud (Linux) does not.
    return hasattr(os, 'startfile')

# --- Google Sheets Logic (Cloud) ---
def get_gsheet_client():
    """Initializes gspread client using Streamlit secrets."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        if "gsheets" not in st.secrets:
            return None, "Secrets 'gsheets' not found"
        
        # Streamlit secrets can be accessed as a dict
        creds_dict = dict(st.secrets["gsheets"])
        # Remove spreadsheet_url from credentials dict before passing to Credentials
        spreadsheet_url = creds_dict.pop("spreadsheet_url", None)
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        return client, spreadsheet_url
    except Exception as e:
        return None, str(e)

def log_to_gsheet(user, event_type, details):
    """Appends a log entry to Google Sheets."""
    client, sheet_url = get_gsheet_client()
    if not client or not sheet_url:
        return
    
    try:
        spreadsheet = client.open_by_url(sheet_url)
        worksheet = spreadsheet.get_worksheet(0)
        
        vn_tz = timezone(timedelta(hours=7))
        timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
        
        worksheet.append_row([timestamp, user, event_type, details])
    except Exception as e:
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
