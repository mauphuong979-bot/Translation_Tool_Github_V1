"""User accounts (username, password, role, auto-fill), stored in a "Users" worksheet of the logging spreadsheet.

The same sheet is used on Windows and on Streamlit Cloud, so a password changed on one is valid on the other.
The first time, the worksheet is created and filled from users.json. When the sheet cannot be reached,
sign-in falls back to users.json (read-only) so the app is not locked; changes then fail with an error.
users.json is rewritten from the sheet whenever they differ, so the fallback stays current.
Passwords are kept as plain text: share the spreadsheet only with admins and the service account.
"""
import json
import os

import streamlit as st

import session_store
import usage_logger


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_FILE = os.path.join(BASE_DIR, "users.json")
SHEET_TITLE = "Users"
HEADERS = ["username", "password", "role", "auto_fill"]


def _local_users():
    if not os.path.exists(LOCAL_FILE):
        return []
    with open(LOCAL_FILE, "r", encoding="utf-8") as f:
        return json.load(f)["users"]


def _to_row(user):
    return [str(user["username"]), str(user["password"]), str(user.get("role", "user")),
            "TRUE" if user.get("auto_fill") else "FALSE"]


def _worksheet():
    import gspread

    spreadsheet = usage_logger.open_spreadsheet()
    try:
        return spreadsheet.worksheet(SHEET_TITLE)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=SHEET_TITLE, rows=200, cols=len(HEADERS))
        rows = [HEADERS] + [_to_row(u) for u in _local_users()]
        worksheet.update("A1", rows, value_input_option="RAW")
        return worksheet


def _sync_local_backup(users):
    """Keep users.json (the offline fallback) equal to the sheet, so admins only edit in one place."""
    try:
        if users and users != _local_users():
            with open(LOCAL_FILE, "w", encoding="utf-8") as f:
                json.dump({"users": users}, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"users.json backup not updated: {e}")


def _read_sheet():
    users = []
    for row in _worksheet().get_all_values()[1:]:
        row = (row + [""] * len(HEADERS))[:len(HEADERS)]
        if not row[0].strip():
            continue
        users.append({"username": row[0].strip(), "password": row[1], "role": row[2].strip() or "user",
                      "auto_fill": row[3].strip().upper() == "TRUE"})
    _sync_local_backup(users)
    return users


@st.cache_data(ttl=60, show_spinner=False)
def _cached_users():
    return _read_sheet()


def load_users(use_cache=True):
    """All users. Falls back to users.json (and records the error) when the sheet cannot be read."""
    try:
        users = _cached_users() if use_cache else _read_sheet()
        st.session_state.pop("users_store_error", None)
        return users
    except Exception as e:
        st.session_state["users_store_error"] = str(e)
        return _local_users()


def _write(users):
    worksheet = _worksheet()
    old_count = len(worksheet.get_all_values())
    rows = [HEADERS] + [_to_row(u) for u in users]
    if len(rows) > worksheet.row_count:
        worksheet.add_rows(len(rows) - worksheet.row_count + 50)
    worksheet.update("A1", rows, value_input_option="RAW")
    if old_count > len(rows):
        worksheet.batch_clear([f"A{len(rows) + 1}:D{old_count}"])
    _cached_users.clear()


def _change(modify, done_message):
    """Read the latest users from the sheet, apply modify(users) -> error or None, write back."""
    try:
        users = _read_sheet()
    except Exception as e:
        return False, f"Không kết nối được Google Sheet, chưa lưu thay đổi: {e}"
    error = modify(users)
    if error:
        return False, error
    try:
        _write(users)
    except Exception as e:
        return False, f"Lỗi khi lưu lên Google Sheet: {e}"
    return True, done_message


def check_credentials(username, password):
    return any(u["username"] == username and u["password"] == password for u in load_users())


def save_user(username, password, role, auto_fill=False):
    def modify(users):
        if any(u["username"] == username for u in users):
            return f"Tên đăng nhập '{username}' đã tồn tại."
        users.append({"username": username, "password": password, "role": role, "auto_fill": auto_fill})
    return _change(modify, f"Đã thêm người dùng '{username}'.")


def remove_user(username):
    def modify(users):
        users[:] = [u for u in users if u["username"] != username]
    result = _change(modify, f"Đã xóa người dùng '{username}'.")
    if result[0]:
        session_store.revoke_user(username)
    return result


def update_user_data(old_username, new_username, new_password, new_role, auto_fill):
    def modify(users):
        if old_username != new_username and any(u["username"] == new_username for u in users):
            return f"Tên đăng nhập '{new_username}' đã tồn tại."
        user = next((u for u in users if u["username"] == old_username), None)
        if user is None:
            return f"Không tìm thấy người dùng '{old_username}'."
        changed["sign_out"] = user["password"] != new_password or old_username != new_username
        user.update(username=new_username, password=new_password, role=new_role, auto_fill=auto_fill)
    changed = {}
    result = _change(modify, f"Đã cập nhật người dùng '{new_username}'.")
    if result[0] and changed.get("sign_out"):
        session_store.revoke_user(old_username)
    return result


def change_password(username, current_password, new_password):
    """A user changing their own password; the current password must match."""
    def modify(users):
        user = next((u for u in users if u["username"] == username), None)
        if user is None:
            return "Không tìm thấy tài khoản."
        if user["password"] != current_password:
            return "Mật khẩu hiện tại không đúng."
        user["password"] = new_password
    result = _change(modify, "Đã đổi mật khẩu.")
    if result[0]:
        session_store.revoke_user(username)
    return result
