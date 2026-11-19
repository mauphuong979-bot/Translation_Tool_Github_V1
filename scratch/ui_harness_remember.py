"""Dev-only: runs app.py against an in-memory spreadsheet (one test user "tester" / "test-pass-1"), to check
sign-in, "remember me" and sign-out in a browser without touching the real Google Sheet or users.json."""
import os
import runpy
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path[:0] = [ROOT, HERE]
os.chdir(ROOT)

import streamlit as st  # noqa: E402

import usage_logger  # noqa: E402
import user_store  # noqa: E402
from fake_gsheet import FakeSpreadsheet  # noqa: E402


@st.cache_resource
def shared_sheet():
    return FakeSpreadsheet()


TEST_USERS = [{"username": "admin", "password": "admin-pass-1", "role": "admin", "auto_fill": False},
              {"username": "tester", "password": "test-pass-1", "role": "user", "auto_fill": False}]
usage_logger.open_spreadsheet = shared_sheet
user_store._local_users = lambda: TEST_USERS
user_store._sync_local_backup = lambda users: None

class _NoCookieContext:
    """?sim_nocookie=1 imitates a proxy (like Streamlit Community Cloud) that does not pass cookies to the app."""

    def __init__(self, real):
        self._real = real

    cookies = {}

    def __getattr__(self, name):
        return getattr(self._real, name)


if st.query_params.get("sim_nocookie") and not isinstance(st.context, _NoCookieContext):
    st.context = _NoCookieContext(st.context)

runpy.run_path(os.path.join(ROOT, "app.py"), run_name="__main__")
