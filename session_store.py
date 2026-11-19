"""'Remember me' sign-in: a random token kept in a browser cookie, checked against a "Sessions" worksheet.

The sheet stores only a SHA-256 hash of each token (so someone reading the sheet cannot reuse it), the user name
and an expiry date. Deleting a row on the sheet signs that browser out; changing a password removes all rows of
that user. No password is ever stored in the browser.
"""
import hashlib
import secrets
from datetime import datetime, timedelta, timezone

import usage_logger


COOKIE_NAME = "fsp_remember"
REMEMBER_DAYS = 30
SHEET_TITLE = "Sessions"
HEADERS = ["token_hash", "username", "expires_at", "created_at"]
_FORMAT = "%Y-%m-%d %H:%M:%S"
_VN = timezone(timedelta(hours=7))


def _hash(token):
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def _now():
    return datetime.now(_VN)


def _worksheet():
    import gspread

    spreadsheet = usage_logger.open_spreadsheet()
    try:
        return spreadsheet.worksheet(SHEET_TITLE)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=SHEET_TITLE, rows=200, cols=len(HEADERS))
        worksheet.update("A1", [HEADERS], value_input_option="RAW")
        return worksheet


def _expired(text):
    try:
        return datetime.strptime(text, _FORMAT).replace(tzinfo=_VN) < _now()
    except ValueError:
        return True


def _delete_rows(worksheet, predicate):
    """Delete data rows (bottom-up, so row numbers stay valid) for which predicate(row) is true."""
    rows = worksheet.get_all_values()
    for number in range(len(rows), 1, -1):
        row = (rows[number - 1] + [""] * len(HEADERS))[:len(HEADERS)]
        if predicate(row):
            worksheet.delete_rows(number)


def create(username):
    """New token for username (expired rows are cleaned up on the way). Returns the token, or None on error."""
    token = secrets.token_urlsafe(32)
    try:
        worksheet = _worksheet()
        _delete_rows(worksheet, lambda row: _expired(row[2]))
        expires = (_now() + timedelta(days=REMEMBER_DAYS)).strftime(_FORMAT)
        worksheet.append_row([_hash(token), username, expires, _now().strftime(_FORMAT)], value_input_option="RAW")
        return token
    except Exception as e:
        print(f"Remember-me token not saved: {e}")
        return None


def username_for(token):
    """The user a valid, unexpired token belongs to, or None (also when the sheet cannot be reached)."""
    if not token:
        return None
    try:
        token_hash = _hash(token)
        for row in _worksheet().get_all_values()[1:]:
            row = (row + [""] * len(HEADERS))[:len(HEADERS)]
            if row[0] == token_hash and not _expired(row[2]):
                return row[1] or None
    except Exception as e:
        print(f"Remember-me check failed: {e}")
    return None


def revoke(token):
    if not token:
        return
    try:
        token_hash = _hash(token)
        _delete_rows(_worksheet(), lambda row: row[0] == token_hash)
    except Exception as e:
        print(f"Remember-me revoke failed: {e}")


def revoke_user(username):
    """Sign username out of every remembered browser."""
    try:
        _delete_rows(_worksheet(), lambda row: row[1] == username)
    except Exception as e:
        print(f"Remember-me revoke for {username} failed: {e}")


def set_cookie_script(token):
    max_age = REMEMBER_DAYS * 24 * 3600
    return _cookie_script(f"{COOKIE_NAME}={token}; max-age={max_age}")


def clear_cookie_script():
    return _cookie_script(f"{COOKIE_NAME}=; max-age=0")


def _cookie_script(cookie):
    # components.html runs in a same-origin iframe, so the cookie is set on the app's own page.
    return f"""<script>
        var secure = window.parent.location.protocol === 'https:' ? '; Secure' : '';
        window.parent.document.cookie = "{cookie}; path=/; SameSite=Lax" + secure;
    </script>"""
