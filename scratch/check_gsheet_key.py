"""Check the Google Sheets key in .streamlit/secrets.toml without printing any secret.

Run from the project folder:  python scratch/check_gsheet_key.py
It only reports the key's shape, whether it loads, and the spreadsheet's tab names (read-only).
"""
import os
import sys
import tomllib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from cryptography.hazmat.primitives import serialization  # noqa: E402

from usage_logger import normalize_private_key  # noqa: E402

path = os.path.join(ROOT, ".streamlit", "secrets.toml")
print("File:", path)
try:
    with open(path, "rb") as handle:
        secrets = tomllib.load(handle)
except Exception as e:
    sys.exit(f"[LỖI] Không đọc được file TOML: {e}")

section = secrets.get("connections", {}).get("gsheets") or secrets.get("gsheets")
if not section:
    sys.exit("[LỖI] Không có mục [connections.gsheets] trong file.")

key = section.get("private_key", "")
print("\n--- private_key (chỉ mô tả, không in giá trị) ---")
print("Độ dài:", len(key), "ký tự (khóa thật thường ~1.700)")
print("Có '-----BEGIN PRIVATE KEY-----':", "BEGIN PRIVATE KEY" in key)
print("Có '-----END PRIVATE KEY-----':", "END PRIVATE KEY" in key)
print("Số lần xuống dòng thật:", key.count("\n"), "| Số chuỗi chữ \\n:", key.count("\\n"), "| Ký tự \\r:", key.count("\r"))


def loads(text):
    try:
        serialization.load_pem_private_key(text.encode(), None)
        return "đọc được"
    except Exception as e:
        return f"KHÔNG đọc được ({type(e).__name__})"


print("Khóa nguyên bản:", loads(key.replace("\\n", "\n")))
print("Khóa sau khi app tự chuẩn hóa:", loads(normalize_private_key(key)))

print("\n--- Các mục khác ---")
for field in ["type", "project_id", "client_email", "token_uri", "spreadsheet"]:
    value = str(section.get(field, "") or section.get("spreadsheet_url", "") if field == "spreadsheet" else section.get(field, ""))
    print(f"{field}: {'có' if value else 'THIẾU'}")
print("client_email kết thúc bằng .iam.gserviceaccount.com:",
      str(section.get("client_email", "")).endswith(".iam.gserviceaccount.com"))

print("\n--- Thử kết nối Google Sheet (chỉ đọc) ---")
try:
    import gspread
    from google.oauth2.service_account import Credentials

    info = {k: v for k, v in section.items() if k not in ("spreadsheet", "spreadsheet_url")}
    info["private_key"] = normalize_private_key(info.get("private_key", ""))
    creds = Credentials.from_service_account_info(
        info, scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    sheet = gspread.authorize(creds).open_by_url(section.get("spreadsheet") or section.get("spreadsheet_url"))
    print("[OK] Kết nối được. Các tab:", [ws.title for ws in sheet.worksheets()])
except Exception as e:
    print(f"[LỖI] {type(e).__name__}: {e}")
