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
    """New token for username (expired rows are cleaned up on the way). Returns (token, None) or (None, error)."""
    token = secrets.token_urlsafe(32)
    try:
        worksheet = _worksheet()
        _delete_rows(worksheet, lambda row: _expired(row[2]))
        expires = (_now() + timedelta(days=REMEMBER_DAYS)).strftime(_FORMAT)
        worksheet.append_row([_hash(token), username, expires, _now().strftime(_FORMAT)], value_input_option="RAW")
        return token, None
    except Exception as e:
        print(f"Remember-me token not saved: {e}")
        return None, str(e)


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


BRIDGE_PARAM = "rt"


BRIDGE_INPUT_KEY = "fsp_rt"


def bridge_script():
    """Fallback for hosts whose proxy does not pass cookies to the app (e.g. Streamlit Community Cloud).
    The browser reads its own cookie and types the token into a hidden text box (key BRIDGE_INPUT_KEY), which
    makes the app rerun with it, without reloading the page. If that box cannot be found, it reloads the page
    once with the token as a query parameter instead."""
    return f"""<script>
        (function () {{
            var page = window.parent;
            function noRemember() {{ page.document.body.classList.add("fsp-no-remember"); }}
            try {{
                var match = page.document.cookie.match(/(?:^|; ){COOKIE_NAME}=([^;]+)/);
                if (!match) {{ noRemember(); return; }}
                var token = match[1], tries = 0, sends = 0;
                function type(input) {{
                    // React only sees values set through the native setter followed by an input event;
                    // Streamlit applies a text box value on Enter.
                    var setValue = Object.getOwnPropertyDescriptor(page.HTMLInputElement.prototype, "value").set;
                    setValue.call(input, "");
                    input.dispatchEvent(new page.Event("input", {{ bubbles: true }}));
                    setValue.call(input, token);
                    input.dispatchEvent(new page.Event("input", {{ bubbles: true }}));
                    ["keydown", "keypress", "keyup"].forEach(function (kind) {{
                        input.dispatchEvent(new page.KeyboardEvent(kind, {{ key: "Enter", code: "Enter",
                            keyCode: 13, which: 13, charCode: 13, bubbles: true }}));
                    }});
                }}
                (function send() {{
                    var input = page.document.querySelector(".st-key-{BRIDGE_INPUT_KEY} input");
                    if (input) {{
                        setTimeout(function () {{
                            type(input);
                            // Still on this page after 3 s (the first try was not picked up)? Try once more.
                            if (++sends < 2) setTimeout(function () {{
                                if (input.isConnected) send();
                            }}, 3000);
                        }}, 300);
                    }} else if (++tries < 30) {{
                        setTimeout(send, 100);
                    }} else {{
                        var url = new URL(page.location.href);
                        url.searchParams.set("{BRIDGE_PARAM}", token);
                        // The component iframe may not navigate the page itself; a script in the page can.
                        var run = page.document.createElement("script");
                        run.textContent = "window.location.replace(" + JSON.stringify(url.toString()) + ");";
                        page.document.body.appendChild(run);
                    }}
                }})();
            }} catch (e) {{
                console.log("remember-me bridge:", e);
                try {{ noRemember(); }} catch (ignored) {{}}
            }}
        }})();
    </script>"""


def debug_script():
    """Shows, from the browser side, whether the remember-me cookie exists (never its value)."""
    return f"""<div style="font:13px sans-serif;color:#555" id="d">...</div><script>
        var d = document.getElementById("d");
        try {{
            var has = window.parent.document.cookie.indexOf("{COOKIE_NAME}=") >= 0;
            d.textContent = "Trình duyệt có cookie ghi nhớ: " + (has ? "CÓ" : "KHÔNG")
                + " | Trang: " + window.parent.location.host + window.parent.location.pathname;
        }} catch (e) {{ d.textContent = "Không đọc được cookie từ trình duyệt: " + e; }}
    </script>"""
