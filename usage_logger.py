import csv
import os
from datetime import datetime, timedelta, timezone

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def is_local_env():
    """Detects if the app is running locally (Windows) or on Streamlit Cloud (Linux)."""
    # Local host (Windows) has os.startfile, Streamlit Cloud (Linux) does not.
    return hasattr(os, 'startfile')

def log_event(user, event_type, details):
    """Logs an event to the usage_log.csv file ONLY IF running locally."""
    if not is_local_env():
        # Future: Call log_event_cloud(user, event_type, details) here
        return

    # Vietnam Time (GMT+7)
    vn_tz = timezone(timedelta(hours=7))
    timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
    file_exists = os.path.isfile(LOG_FILE)
    
    try:
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Write header if file is new
            if not file_exists:
                writer.writerow(["Timestamp", "User", "Event Type", "Details"])
            
            writer.writerow([timestamp, user, event_type, details])
    except Exception as e:
        # Fail silently in the UI but print for debugging
        print(f"Error logging event: {e}")

def get_logs():
    """Reads all logs from the CSV file."""
    if not os.path.exists(LOG_FILE):
        return []
    
    try:
        logs = []
        with open(LOG_FILE, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                logs.append(row)
        return logs[::-1]  # Return in reverse chronological order
    except Exception:
        return []
