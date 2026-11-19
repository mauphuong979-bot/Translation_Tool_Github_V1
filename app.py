import streamlit as st # v1.1 - Fixed case_threshold sync
import pandas as pd
import io
import json
import os
import base64
import re
import streamlit.components.v1 as components
import translation_lib as tl
from processor import process_financial_report
from review_export import build_ai_review_document
from review_import_ui import render_ai_import
from review_filenames import review_output_filename
from usage_logger import log_event, get_logs, is_local_env
import session_store
from user_store import load_users, check_credentials, save_user, remove_user, update_user_data, change_password
from datetime import datetime, timedelta, timezone
import metadata_extractor as mex
from user_guide import show_user_guide
from review_annotations import needs_ai_translation
from docx import Document
import time


# Display names (stored values stay unchanged)
LANGUAGE_NAMES = {"E": "E – Tiếng Anh", "Hs": "Hs – Tiếng Hoa giản thể", "Ht": "Ht – Tiếng Hoa phồn thể"}
OUTPUT_MODE_NAMES = {
    "Original output (current)": "Bản gốc (gợi ý chèn trong văn bản)",
    "Clean output (no inline suggestions)": "Bản sạch (không chèn gợi ý)",
    "Comments + Track Changes": "Comments + Track Changes (khuyên dùng)",
    "All outputs": "Tất cả các bản",
}

# Page Configuration
st.set_page_config(
    page_title="Financial Statements Processor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load CSS
with open("style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# --- Helpers ---
def st_auto_download(filedata, filename, mime):
    """Triggers a browser download immediately using JavaScript."""
    # Convert BytesIO to actual bytes if necessary
    if hasattr(filedata, 'getvalue'):
        data = filedata.getvalue()
    else:
        data = filedata
        
    b64 = base64.b64encode(data).decode()
    js = f"""
        <script>
            var link = document.createElement('a');
            link.href = 'data:{mime};base64,{b64}';
            link.download = '{filename}';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        </script>
    """
    components.html(js, height=0)

def summarize_outputs(outputs, review_items, seconds):
    """Quick counts shown after processing: Track Changes, comments and items left for AI."""
    stats = [("Thời gian xử lý", f"{seconds:.0f}s")]
    if outputs.get("review"):
        document = Document(io.BytesIO(outputs["review"]))
        insertions = len(document.element.xpath(".//w:ins"))
        stats += [("Track Changes", insertions), ("Comment gợi ý", len(document.comments)),
                  ("Mục cần AI", sum(1 for item in review_items or [] if needs_ai_translation(item)))]
    else:
        stats.append(("File kết quả", len(outputs)))
    return stats


def simple_date_format(d_str):
    """
    Converts Vietnamese string dates (e.g., '31 tháng 12 năm 2025') to 'dd/mm/yyyy'.
    """
    if not d_str or not isinstance(d_str, str): 
        return d_str
    # Match: day tháng month năm year
    match = re.search(r"(\d{1,2})\s+th\u00e1ng\s+(\d{1,2})\s+n\u0103m\s+(\d{4})", d_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"
    return d_str

def highlight_match(text, keyword):
    """Wraps matching keyword in a span for blue coloring."""
    if not keyword or not isinstance(text, str):
        return text
    pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    return pattern.sub(lambda m: f'<span class="dict-highlight">{m.group(0)}</span>', text)

# Authentication Logic
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Users are stored in the "Users" worksheet of the logging Google Sheet (see user_store.py)

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False


def sign_in(username, how):
    st.session_state.authenticated = True
    st.session_state.username = username
    log_event(username, "Login", how)


# Cookie changes queued before a rerun (set on "remember me", cleared on sign-out) run here, on the next page draw.
if "cookie_script" in st.session_state:
    components.html(st.session_state.pop("cookie_script"), height=0)

def waiting_overlay(text, auto_hide=False):
    """Full-page spinner shown while a remembered sign-in is checked, so the login form does not flash.
    auto_hide: the browser-side bridge hides it when there is no cookie; it also disappears by itself after
    8 seconds in case that script never runs, so nobody is left in front of a spinner."""
    hide = ("body.fsp-no-remember .fsp-overlay { display: none; }"
            ".fsp-overlay { animation: fsp-giveup 0s 8s forwards; }"
            "@keyframes fsp-giveup { to { visibility: hidden; } }") if auto_hide else ""
    return f"""<style>
        .fsp-overlay {{ position: fixed; inset: 0; z-index: 999990; background: #fff; display: flex;
            flex-direction: column; align-items: center; justify-content: center; gap: 14px;
            font-family: sans-serif; color: #555; font-size: 15px; }}
        .fsp-spin {{ width: 38px; height: 38px; border-radius: 50%; border: 4px solid #e6e6e6;
            border-top-color: #ff4b4b; animation: fsp-rotate 0.8s linear infinite; }}
        @keyframes fsp-rotate {{ to {{ transform: rotate(360deg); }} }}
        {hide}
    </style><div class="fsp-overlay"><div class="fsp-spin"></div><div>{text}</div></div>"""


# "Remember me": a valid token in this browser's cookie signs the user in without the password (once per visit).
# The token normally arrives with the cookie; where the host's proxy drops cookies (Streamlit Community Cloud),
# the login screen's bridge script reloads the page once with the token in the "rt" query parameter instead.
server_cookie_token = st.context.cookies.get(session_store.COOKIE_NAME)
bridge_token = st.query_params.get(session_store.BRIDGE_PARAM)
if bridge_token:
    del st.query_params[session_store.BRIDGE_PARAM]
    st.session_state.bridge_used = True
if not st.session_state.authenticated and not st.session_state.get("auto_login_checked"):
    st.session_state.auto_login_checked = True
    remembered_token = server_cookie_token or bridge_token
    if remembered_token:
        st.session_state.remember_source = "cookie" if server_cookie_token else "bridge"
        overlay = st.empty()
        overlay.markdown(waiting_overlay("Đang đăng nhập…"), unsafe_allow_html=True)
        remembered_user = session_store.username_for(remembered_token)
        valid = remembered_user and any(u["username"] == remembered_user for u in load_users())
        overlay.empty()
        if valid:
            st.session_state.remember_token = remembered_token
            sign_in(remembered_user, f"Signed in from a remembered browser ({st.session_state.remember_source})")
        else:
            # Expired or revoked: forget it in the browser so the bridge does not try it again.
            components.html(session_store.clear_cookie_script(), height=0)

if st.query_params.get("debug_login"):
    with st.expander("🔎 Chẩn đoán ghi nhớ đăng nhập", expanded=True):
        st.write({
            "Server nhận được cookie": bool(server_cookie_token),
            "Đã dùng đường dự phòng (rt)": bool(st.session_state.get("bridge_used")),
            "Nguồn token lần này": st.session_state.get("remember_source"),
            "Đang đăng nhập": st.session_state.authenticated,
            "Lỗi lưu ghi nhớ gần nhất": st.session_state.get("remember_error"),
        })
        components.html(session_store.debug_script(), height=40)

def handle_name_cap_change():
    """Callback to automatically convert Translated Name to Proper Case if all caps."""
    if 'meta_name_cap' in st.session_state:
        st.session_state.meta_name_cap = tl.ensure_proper_case(st.session_state.meta_name_cap)

def login_screen():
    _, login_column, _ = st.columns([1, 1.1, 1])
    with login_column, st.container(border=True, key="login_box"):
        login_form()


def login_form():
    st.markdown('<div class="login-header">📄 Financial Statements Processor</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-sub">🔐 Đăng nhập để tiếp tục</div>', unsafe_allow_html=True)

    users = load_users()
    usernames = [u["username"] for u in users]
    if st.session_state.get("users_store_error"):
        st.warning("⚠️ Không kết nối được Google Sheet, đang dùng danh sách tài khoản dự phòng trên máy "
                   "(mật khẩu đã đổi gần đây có thể chưa có hiệu lực).")
        with st.expander("Chi tiết lỗi"):
            st.code(st.session_state.users_store_error, language=None, wrap_lines=True)

    with st.form("login_form"):
        username = st.selectbox("Tên đăng nhập", usernames, index=None, placeholder="Chọn tên đăng nhập")
        password = st.text_input("Mật khẩu", type="password")
        remember = st.checkbox("Ghi nhớ đăng nhập trên máy này", value=True,
                               help=f"Không cần nhập mật khẩu trong {session_store.REMEMBER_DAYS} ngày trên trình duyệt này. "
                                    "Không dùng trên máy dùng chung.")
        submit = st.form_submit_button("Đăng nhập", type="primary", use_container_width=True)

        if submit:
            if not username:
                st.error("Vui lòng chọn tên đăng nhập.")
            elif check_credentials(username, password):
                sign_in(username, "Successfully signed in")
                if remember:
                    token, error = session_store.create(username)
                    if token:
                        st.session_state.remember_token = token
                        st.session_state.cookie_script = session_store.set_cookie_script(token)
                    else:
                        st.session_state.remember_error = error
                st.rerun()
            else:
                st.error("Sai tên đăng nhập hoặc mật khẩu.")

if not st.session_state.authenticated:
    # Bridge: once per visit, unless this visit already came through it or the user just signed out.
    if not server_cookie_token and not st.session_state.get("bridge_used") and not st.session_state.get("signed_out"):
        if not st.session_state.get("bridge_checked"):
            # First draw of this visit: cover the form until the browser says whether a remembered sign-in exists.
            st.markdown(waiting_overlay("Đang kiểm tra đăng nhập…", auto_hide=True), unsafe_allow_html=True)
            st.session_state.bridge_checked = True
        components.html(session_store.bridge_script(), height=0)
    login_screen()
    st.stop()

if st.session_state.get("remember_error"):
    st.warning("⚠️ Đã đăng nhập nhưng chưa lưu được ghi nhớ đăng nhập (lần sau vẫn phải nhập mật khẩu): "
               f"{st.session_state.pop('remember_error')}")

# Initialize Settings
if 'processed_file_id' not in st.session_state:
    st.session_state.processed_file_id = None
if 'processed_output_word' not in st.session_state:
    st.session_state.processed_output_word = None
if 'processed_outputs' not in st.session_state:
    st.session_state.processed_outputs = {}
if 'processed_output_excel' not in st.session_state:
    st.session_state.processed_output_excel = None
if 'process_success_msg' not in st.session_state:
    st.session_state.process_success_msg = None
if 'processed_filename' not in st.session_state:
    st.session_state.processed_filename = ""
if 'auto_open' not in st.session_state:
    st.session_state.auto_open = True
if 'case_threshold' not in st.session_state:
    st.session_state.case_threshold = 30

# App UI (Authenticated Only)
# Initialize Metadata Fields in Session State with defaults if they don't exist
metadata_defaults = {
    "meta_name_lc": "ABC Company",
    "meta_name_cap": "CustomerName",
    "meta_year_end": "",
    "meta_date": "",
    "meta_period_short": "",
    "meta_period_short_2": "",
    "meta_signer_1": "",
    "meta_signer_2": "",
    "meta_signer_3": "",
    "meta_translate": "E"
}

for key, default_val in metadata_defaults.items():
    # Re-assigning keeps the value even when an upload triggers st.rerun() before the
    # metadata widgets are drawn; otherwise Streamlit drops the widget state and the
    # field silently falls back to its default.
    st.session_state[key] = st.session_state.get(key, default_val)
st.session_state.export_mode = st.session_state.get("export_mode", "Comments + Track Changes")



PROCESS_STEPS = [
    ("unicode", "1. Chuẩn hóa Unicode (NFC)"),
    ("clean_v", "2. Sửa lỗi theo CleanV"),
    ("para_template", "3. Mẫu đoạn văn"),
    ("dictionary", "4. Dịch theo từ điển"),
    ("dual_font", "5. Font kép tiếng Hoa"),
    ("number_swap", "6. Đổi định dạng số"),
    ("table_size", "7. Bảng & bố cục"),
    ("table_7pt", "7.5. Bảng >10 cột thu về cỡ 7pt"),
    ("date_format", "8. Định dạng ngày"),
    ("textbox", "9. Textbox & chữ Draft"),
    ("signer_accents", "10. Dấu tên người ký"),
    ("highlight", "11. Tô vàng tiếng Việt chưa dịch"),
    ("suggestion", "12. Gợi ý từ điển"),
    ("always_original", "13. Luôn hiện [Original]"),
]
if 'process_steps' not in st.session_state:
    st.session_state.process_steps = {key: True for key, _ in PROCESS_STEPS}
for key, _ in PROCESS_STEPS:
    st.session_state.setdefault(f"step_{key}", st.session_state.process_steps.get(key, True))


def set_all_steps(value):
    for key, _ in PROCESS_STEPS:
        st.session_state[f"step_{key}"] = value


# Top bar: compact title on the left, help / settings / account on the right
with st.container(key="top_bar", horizontal=True, vertical_alignment="center"):
    st.markdown('<div class="app-title">📄 Financial Statements Processor</div>', unsafe_allow_html=True)
    st.space("stretch")
    if st.button("❓", key="open_guide", help="Hướng dẫn sử dụng"):
        show_user_guide()
    with st.popover("⚙️", help="Các bước xử lý và tùy chọn"):
        st.markdown("**Các bước xử lý**")
        all_col, none_col = st.columns(2)
        all_col.button("Chọn tất cả", key="steps_all", on_click=set_all_steps, args=(True,), use_container_width=True)
        none_col.button("Bỏ chọn tất cả", key="steps_none", on_click=set_all_steps, args=(False,), use_container_width=True)
        step_columns = st.columns(2)
        for index, (key, label) in enumerate(PROCESS_STEPS):
            with step_columns[0 if index < (len(PROCESS_STEPS) + 1) // 2 else 1]:
                st.checkbox(label, key=f"step_{key}")
        st.divider()
        st.markdown("**Tùy chọn**")
        st.checkbox(
            "Tự mở/tải kết quả",
            key="auto_open",
            help="Xử lý xong sẽ tự mở file bằng Word (máy Windows) hoặc tự tải xuống (bản cloud/Linux)."
        )
    role = "Quản trị viên" if st.session_state.username == "admin" else "Người dùng"
    with st.popover(f"👤 {st.session_state.username}"):
        st.markdown(f"**{st.session_state.username}**  \n{role}")
        with st.expander("🔑 Đổi mật khẩu"):
            with st.form("change_password_form", clear_on_submit=True):
                current_pass = st.text_input("Mật khẩu hiện tại", type="password")
                new_pass_1 = st.text_input("Mật khẩu mới", type="password")
                new_pass_2 = st.text_input("Nhập lại mật khẩu mới", type="password")
                if st.form_submit_button("Lưu mật khẩu mới", use_container_width=True):
                    if not current_pass or not new_pass_1:
                        st.warning("Vui lòng nhập mật khẩu hiện tại và mật khẩu mới.")
                    elif new_pass_1 != new_pass_2:
                        st.error("Hai lần nhập mật khẩu mới không khớp.")
                    elif new_pass_1 == current_pass:
                        st.warning("Mật khẩu mới phải khác mật khẩu hiện tại.")
                    else:
                        success, msg = change_password(st.session_state.username, current_pass, new_pass_1)
                        if success:
                            st.success(msg)
                            log_event(st.session_state.username, "User Management", "Changed own password")
                        else:
                            st.error(msg)
        if st.button("🚪 Đăng xuất", use_container_width=True):
            session_store.revoke(st.session_state.pop("remember_token", None)
                                 or st.context.cookies.get(session_store.COOKIE_NAME))
            st.session_state.cookie_script = session_store.clear_cookie_script()
            st.session_state.signed_out = True
            st.session_state.authenticated = False
            st.rerun()

for key, _ in PROCESS_STEPS:
    st.session_state.process_steps[key] = st.session_state[f"step_{key}"]

# Main Interface Tabs Setup
tab_titles = ["🚀 Xử lý", "📖 Tra từ điển"]
if st.session_state.authenticated and st.session_state.username == "admin":
    tab_titles.append("📈 Quản trị")

tabs = st.tabs(tab_titles)

with tabs[0]:
    # --- Tab 1: Process ---
    # 1. Report to translate
    with st.container(border=True, key="report_box"):
        st.markdown('<div class="section-title">📂 Báo cáo cần dịch <span>file Word .docx</span></div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Chọn file báo cáo (.docx)", type=["docx"], key="report_file", label_visibility="collapsed")

    # Automatic Metadata Extraction
    if uploaded_file:
        # Use a session state marker to ensure extraction happens only once per new file
        if st.session_state.get('last_extracted_id') != uploaded_file.name:
            # Reset processed data for the new file
            st.session_state.processed_file_id = None
            st.session_state.processed_output_word = None
            st.session_state.processed_outputs = {}
            st.session_state.processed_output_excel = None
            st.session_state.process_success_msg = None
            st.session_state.process_stats = None

            with st.spinner("Đang trích xuất thông tin báo cáo..."):
                # We need to seek(0) in case the file pointer was moved
                uploaded_file.seek(0)
                ext_meta = mex.extract_metadata(uploaded_file)
                uploaded_file.seek(0) # Reset again for later processing

                # Update session state fields
                if ext_meta.get("name_vn"):
                    st.session_state.meta_name_lc = ext_meta["name_vn"]
                if ext_meta.get("year_end"):
                    st.session_state.meta_year_end = ext_meta["year_end"]
                if ext_meta.get("report_date"):
                    st.session_state.meta_date = ext_meta["report_date"]
                if ext_meta.get("period_in"):
                    st.session_state.meta_period_short = ext_meta["period_in"]
                if ext_meta.get("period_in_2"):
                    st.session_state.meta_period_short_2 = ext_meta["period_in_2"]
                if ext_meta.get("signer_1"):
                    st.session_state.meta_signer_1 = ext_meta["signer_1"]
                if ext_meta.get("signer_2"):
                    st.session_state.meta_signer_2 = ext_meta["signer_2"]
                if ext_meta.get("signer_3"):
                    st.session_state.meta_signer_3 = ext_meta["signer_3"]

                st.session_state.last_extracted_id = uploaded_file.name
                st.rerun() # Refresh to show extracted values in inputs

    # 2. Prior year reports (compact)
    def prior_year_summary(meta, filename):
        if not meta:
            return f"✓ {filename}"
        tax_code = re.sub(r"^[^:]*:\s*", "", meta.get("tax_code", "") or "")
        details = " · ".join(value for value in (meta.get("company_name", ""), f"MST {tax_code}" if tax_code else "",
                                                 meta.get("reporting_period", "")) if value)
        return f"✓ **{filename}**  \n{details}"

    def prior_year_uploader(label, uploader_key, prefix, spinner_text):
        uploaded = st.file_uploader(label, type=["docx"], key=uploader_key)
        if uploaded:
            if st.session_state.get(f"{prefix}_last_filename") != uploaded.name:
                with st.spinner(spinner_text):
                    uploaded.seek(0)
                    st.session_state[f"{prefix}_meta"] = mex.extract_first_table_metadata(uploaded)
                    st.session_state[f"{prefix}_last_filename"] = uploaded.name
                    company = (st.session_state[f"{prefix}_meta"] or {}).get("company_name", "")
                    if prefix == "py_trans" and company and st.session_state.meta_name_cap in ("", metadata_defaults["meta_name_cap"]):
                        st.session_state.meta_name_cap = tl.ensure_proper_case(company)
                    st.rerun()
            st.caption(prior_year_summary(st.session_state.get(f"{prefix}_meta"), uploaded.name))
        else:
            st.session_state.pop(f"{prefix}_last_filename", None)
            st.session_state.pop(f"{prefix}_meta", None)
        return uploaded

    py_dict = None
    with st.container(border=True, key="py_box"):
        title_column, toggle_column = st.columns([3, 2], vertical_alignment="center")
        with title_column:
            st.markdown('<div class="section-title">📅 Báo cáo năm trước <span>tùy chọn</span></div>', unsafe_allow_html=True)
        with toggle_column:
            suggest_py = st.toggle(
                "Dùng báo cáo năm trước để gợi ý",
                value=st.session_state.get("suggest_py", True),
                key="suggest_py"
            )
        if suggest_py:
            col_py_1, col_py_2 = st.columns(2)
            with col_py_1:
                py_vn_file = prior_year_uploader("Bản tiếng Việt năm trước (.docx)", "py_vn_uploader", "py_vn",
                                                 "Đang đọc báo cáo tiếng Việt năm trước...")
            with col_py_2:
                py_trans_file = prior_year_uploader("Bản đã dịch năm trước (.docx)", "py_trans_uploader", "py_trans",
                                                    "Đang đọc báo cáo đã dịch năm trước...")

            # Build py_dict if both files are uploaded
            if py_vn_file and py_trans_file:
                cache_key = f"py_dict_{py_vn_file.name}_{py_trans_file.name}"
                if st.session_state.get("py_dict_cache_key") != cache_key:
                    with st.spinner("Đang tạo từ điển từ báo cáo năm trước..."):
                        py_dict = tl.build_py_dictionary(py_vn_file, py_trans_file)
                        st.session_state.py_dict = py_dict
                        st.session_state.py_dict_cache_key = cache_key
                else:
                    py_dict = st.session_state.get("py_dict")
            else:
                st.session_state.pop("py_dict", None)
                st.session_state.pop("py_dict_cache_key", None)
        else:
            # Clear all cache when suggest_py is unchecked
            for state_key in ("py_vn_last_filename", "py_vn_meta", "py_trans_last_filename", "py_trans_meta", "py_dict", "py_dict_cache_key"):
                st.session_state.pop(state_key, None)

    # 3. Report metadata
    with st.expander("📝 **Thông tin báo cáo**", expanded=True):
        col_metadata_1, col_metadata_2 = st.columns(2)

        with col_metadata_1:
            # 1. Translate into
            lang_options = ["E", "Hs", "Ht"]
            if st.session_state.meta_translate not in lang_options:
                st.session_state.meta_translate = "E"
            st.selectbox("Dịch sang", options=lang_options, key="meta_translate", format_func=LANGUAGE_NAMES.get)

            # 2. Translated Name (not capitalized)
            st.text_input("Tên công ty đã dịch (không viết hoa toàn bộ)", placeholder="Tên công ty bằng ngôn ngữ đích", key="meta_name_cap", on_change=handle_name_cap_change)

            # 3. Name in Vietnamese (not capitalized)
            st.text_input("Tên công ty tiếng Việt (không viết hoa toàn bộ)", placeholder="Công ty TNHH ABC...", key="meta_name_lc")

            # 8. Signer Names (extracted from last table)
            c_s1, c_s2, c_s3 = st.columns(3)
            with c_s1:
                st.text_input("Người ký 1", key="meta_signer_1")
            with c_s2:
                st.text_input("Người ký 2", key="meta_signer_2")
            with c_s3:
                st.text_input("Người ký 3", key="meta_signer_3")

        with col_metadata_2:
            date_1, date_2 = st.columns(2)
            with date_1:
                # 4. Year-end date
                st.text_input("Ngày kết thúc năm tài chính", placeholder="31/12/2025", key="meta_year_end")
            with date_2:
                # 5. Reporting date
                st.text_input("Ngày lập báo cáo", placeholder="08/04/2026", key="meta_date")

            # 6. Period 1 (in table)
            st.text_input("Kỳ báo cáo 1 (trong bảng)", placeholder="Từ xx/xx/20xx đến xx/xx/20xx", key="meta_period_short")

            # 7. Period 2 (in table)
            st.text_input("Kỳ báo cáo 2 (trong bảng)", placeholder="Từ xx/xx/20xx đến xx/xx/20xx", key="meta_period_short_2")

    # 4. Output mode + Process
    if not uploaded_file:
        st.caption("⬆️ Tải báo cáo cần dịch lên để bắt đầu xử lý.")
    if uploaded_file:
        with st.container(border=True, key="action_box"):
            mode_column, button_column = st.columns([3, 1], vertical_alignment="bottom")
            with mode_column:
                export_mode = st.selectbox(
                    "Chế độ xuất file",
                    ["Original output (current)", "Clean output (no inline suggestions)", "Comments + Track Changes", "All outputs"],
                    key="export_mode",
                    format_func=OUTPUT_MODE_NAMES.get,
                    help="Comments + Track Changes: câu khớp 100% với báo cáo năm trước được thay bằng Track Changes, câu chưa khớp có comment gợi ý, kèm bảng AI_Form.docx cho các mục dưới 100%. Tất cả các bản cũng tạo bảng này."
                )
            with button_column:
                process_btn = st.button("🚀 Xử lý báo cáo", use_container_width=True, type="primary")
            st.caption(f"📄 {uploaded_file.name}")

        # Manual button click
        trigger_processing = process_btn

        if trigger_processing:
            with st.spinner("Đang xử lý báo cáo..."):
                started_at = time.perf_counter()
                try:
                    target_col = st.session_state.meta_translate
                    
                    # 0. Load Dynamic Dictionary_v3
                    # Prepare metadata for tag substitution
                    metadata_for_tags = {
                        "name_vn": st.session_state.meta_name_lc,
                        "name_trans": st.session_state.meta_name_cap,
                        "year_end": st.session_state.meta_year_end,
                        "report_date": st.session_state.meta_date,
                        "period_in": st.session_state.meta_period_short,
                        "period_in_2": st.session_state.meta_period_short_2,
                        "signer_1": st.session_state.meta_signer_1,
                        "signer_2": st.session_state.meta_signer_2,
                        "signer_3": st.session_state.meta_signer_3
                    }
                    
                    v3_df = tl.load_and_fill_v3_dictionary(metadata_for_tags)
                    if v3_df is not None:
                        # Filter and prepare translation map for the specific target language
                        if target_col in v3_df.columns:
                            display_df = v3_df[['Vietnamese', target_col]].dropna().copy()
                            st.session_state.current_dict = display_df
                        else:
                            st.error(f"Không tìm thấy cột ngôn ngữ '{target_col}' trong Dictionary_v3.")
                            st.stop()
                    else:
                        st.error("Không đọc được Dictionary_v3.xlsx.")
                        st.stop()

                    # Translation map is now built from the dynamically filled dictionary
                    translation_map = dict(zip(st.session_state.current_dict['Vietnamese'], st.session_state.current_dict[target_col]))
                    # Company names confirmed by users (shared table); dictionary entries keep priority
                    try:
                        from company_names_store import company_name_map
                        for company_vn, company_international in company_name_map().items():
                            translation_map.setdefault(company_vn, company_international)
                    except Exception as company_error:
                        st.warning(f"Không đọc được bảng tên công ty đã xác nhận: {company_error}")

                    # Prepare metadata
                    metadata = {
                        "Name (not capitalized)": tl.clean_text(st.session_state.meta_name_lc),
                        "Reporting date": tl.clean_text(st.session_state.meta_date),
                        "Translate into": st.session_state.meta_translate,
                        "Year-end date": tl.clean_text(st.session_state.meta_year_end),
                        "Translated Name": tl.clean_text(st.session_state.meta_name_cap),
                        "Period (in table)": tl.clean_text(st.session_state.meta_period_short),
                        "signer_1": tl.clean_text(st.session_state.meta_signer_1),
                        "signer_2": tl.clean_text(st.session_state.meta_signer_2),
                        "signer_3": tl.clean_text(st.session_state.meta_signer_3)
                    }
                    
                    # Inject suggest_py and py_dict
                    process_settings = dict(st.session_state.process_steps)
                    process_settings["suggest_py"] = suggest_py
                    active_py_dict = st.session_state.get("py_dict") if suggest_py else None
                    
                    mode_specs = {
                        "Original output (current)": [("original", process_settings)],
                        "Clean output (no inline suggestions)": [("clean", dict(process_settings))],
                        "Comments + Track Changes": [("review", dict(process_settings))],
                        "All outputs": [("original", process_settings), ("clean", dict(process_settings)), ("review", dict(process_settings))],
                    }
                    outputs = {}
                    msg = "Report processed."
                    for variant_name, variant_settings in mode_specs[export_mode]:
                        uploaded_file.seek(0)
                        if variant_name in ("clean", "review"):
                            variant_settings["suggestion"] = False
                            variant_settings["always_original"] = False
                            variant_settings["suggest_py"] = False
                        variant_settings["review_annotations"] = variant_name == "review"
                        review_items = [] if variant_name == "review" else None
                        variant_file, variant_msg = process_financial_report(
                            uploaded_file,
                            metadata=metadata,
                            translation_map=translation_map,
                            case_threshold=st.session_state.case_threshold,
                            target_col=target_col,
                            process_settings=variant_settings,
                            py_dict=active_py_dict,
                            review_items=review_items
                        )
                        if not variant_file:
                            processed_file, msg = None, variant_msg
                            break
                        outputs[variant_name] = variant_file.getvalue()
                        if variant_name == "review":
                            ai_review_file = build_ai_review_document(review_items, target_col)
                            if ai_review_file is not None:
                                outputs["ai_review"] = ai_review_file.getvalue()
                    processed_file = outputs.get("original") or outputs.get("clean") or outputs.get("review")
                    if processed_file:
                        # Prepare the resolved dictionary for download
                        resolved_dict_output = io.BytesIO()
                        with pd.ExcelWriter(resolved_dict_output, engine='openpyxl') as writer:
                            v3_df.to_excel(writer, index=False, sheet_name='ResolvedDictionary')
                        resolved_dict_data = resolved_dict_output.getvalue()

                        # Generate Dynamic Filename
                        original_name = uploaded_file.name.rsplit('.', 1)[0]
                        target_lang = st.session_state.meta_translate
                        # Vietnam Time (GMT+7)
                        vn_tz = timezone(timedelta(hours=7))
                        timestamp = datetime.now(vn_tz).strftime("%d%m%y_%H%M")
                        processed_filename = f"{original_name}_{target_lang}_{timestamp}.docx"
                        primary_output_kind = next(output_kind for output_kind in ("original", "clean", "review") if outputs.get(output_kind))
                        primary_filename = review_output_filename(processed_filename, primary_output_kind)
                        
                        # Save to session state for persistent display
                        st.session_state.processed_file_id = uploaded_file.name
                        st.session_state.processed_output_word = processed_file
                        st.session_state.processed_outputs = outputs
                        st.session_state.processed_output_excel = resolved_dict_data
                        st.session_state.process_success_msg = msg
                        st.session_state.processed_filename = processed_filename
                        st.session_state.process_stats = summarize_outputs(outputs, review_items, time.perf_counter() - started_at)
                        
                        log_event(st.session_state.username, "Processing", f"File: {uploaded_file.name} (Lang: {target_col})")
                        
                        # AUTO OPEN / DOWNLOAD IF ENABLED
                        if st.session_state.auto_open:
                            if hasattr(os, 'startfile'):
                                # Local Windows Usage: Open Word directly
                                temp_dir = os.path.join(BASE_DIR, "temp_output")
                                if not os.path.exists(temp_dir):
                                    os.makedirs(temp_dir)
                                temp_path = os.path.join(temp_dir, primary_filename)
                                try:
                                    with open(temp_path, "wb") as f:
                                        # Handle both BytesIO and raw bytes
                                        f.write(processed_file.getvalue() if hasattr(processed_file, "getvalue") else processed_file)
                                    os.startfile(temp_path)
                                except Exception as e:
                                    st.warning(f"Không tự mở được Word: {e}. Bạn vẫn có thể tải file bằng nút bên dưới.")
                            else:
                                # Cloud / Linux Usage: Trigger automatic browser download
                                st_auto_download(processed_file, primary_filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    else:
                        st.error(msg)
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi khi xử lý: {str(e)}")

        # PERSISTENT DOWNLOAD BUTTONS - Shown if processing was successful for current file
        if st.session_state.get('processed_file_id') == uploaded_file.name:
            # Note: Specific success message removed as requested.
            
            output_labels = {
                "original": "📥 Tải bản gốc",
                "clean": "📥 Tải bản sạch",
                "review": "📥 Tải báo cáo Comments + Track Changes",
                "ai_review": "📥 Tải bảng dịch AI (AI_Form)",
            }
            stats = st.session_state.get("process_stats")
            if stats:
                st.markdown('<div class="ai-review-summary">' + ''.join(
                    f'<div class="ai-review-stat"><span>{label}</span><strong>{value}</strong></div>'
                    for label, value in stats) + '</div>', unsafe_allow_html=True)
            output_columns = st.columns(2)
            for output_index, (output_key, output_data) in enumerate(st.session_state.get("processed_outputs", {}).items()):
                output_name = review_output_filename(st.session_state.processed_filename, output_key)
                output_columns[output_index % 2].download_button(
                    label=output_labels[output_key],
                    data=output_data,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"btn_download_{output_key}_persist",
                    use_container_width=True
                )
            if "review" in st.session_state.get("processed_outputs", {}):
                if "ai_review" in st.session_state.processed_outputs:
                    st.caption("Câu khớp 100% với báo cáo năm trước đã được ghi bằng Track Changes. Bảng AI_Form gồm các gợi ý chưa khớp và đoạn tô vàng chưa có tham chiếu, mỗi mục có một mã riêng. Gửi bảng cho AI điền cột cuối, rồi nhập file AI trả về ở mục bên dưới.")
                else:
                    st.caption("Không còn mục nào cần AI dịch. Câu khớp 100% với báo cáo năm trước đã được ghi bằng Track Changes; hãy xem các comment cần rà soát thủ công trong báo cáo.")

    render_ai_import()

with tabs[1]:
    # --- Tab 2: Dictionary Search ---
    st.markdown("### 📖 **Tra từ điển**")
    search_query = st.text_input("🔍 Từ khóa (tiếng Việt hoặc bản dịch)", placeholder="Nhập từ khóa...", key="dict_search")

    # 1. Load and resolve dictionary with current UI metadata
    metadata_for_display = {
        "name_vn": st.session_state.meta_name_lc,
        "name_trans": st.session_state.meta_name_cap,
        "year_end": st.session_state.meta_year_end,
        "report_date": st.session_state.meta_date,
        "period_in": st.session_state.meta_period_short,
        "period_in_2": st.session_state.meta_period_short_2,
        "signer_1": st.session_state.meta_signer_1,
        "signer_2": st.session_state.meta_signer_2,
        "signer_3": st.session_state.meta_signer_3
    }

    # 2. Language Selection UI
    current_lang = st.session_state.get("meta_translate", "E")
    dict_lang_choices = [
        f"Đang chọn ({current_lang})",
        "E",
        "Hs",
        "Ht",
        "Tất cả"
    ]
    
    selected_lang_choice = st.radio(
        "Hiển thị bản dịch:",
        options=dict_lang_choices,
        index=0,
        horizontal=True,
        key="dict_lang_sel"
    )

    try:
        full_v3_df = tl.load_and_fill_v3_dictionary(metadata_for_display)
        if full_v3_df is not None:
            # 3. Determine target columns to show
            if selected_lang_choice == dict_lang_choices[0]: # Current Language
                target_cols = [current_lang]
            elif selected_lang_choice == "E":
                target_cols = ["E"]
            elif selected_lang_choice == "Hs":
                target_cols = ["Hs"]
            elif selected_lang_choice == "Ht":
                target_cols = ["Ht"]
            else: # "Tất cả"
                target_cols = ["E", "Hs", "Ht"]

            # Filter columns that actually exist in the DataFrame
            available_cols = [c for c in target_cols if c in full_v3_df.columns]
            
            if available_cols:
                cols_to_show = ["Vietnamese"] + available_cols
                display_df = full_v3_df[cols_to_show].copy()
                
                # 4. Filter results based on search query (normalize to NFC for robust matching)
                search_query_norm = ""
                if search_query:
                    import unicodedata
                    search_query_norm = unicodedata.normalize('NFC', search_query)
                
                if search_query_norm:
                    # Search in Vietnamese and all selected translation columns
                    mask = display_df["Vietnamese"].str.contains(search_query_norm, case=False, na=False)
                    for col in available_cols:
                        mask |= display_df[col].str.contains(search_query_norm, case=False, na=False)
                    display_df = display_df[mask]
                
                # 5. Sort by Vietnamese text length (shortest first)
                display_df["_len"] = display_df["Vietnamese"].str.len()
                display_df = display_df.sort_values("_len").drop(columns=["_len"])
                
                # 6. Build Styled HTML Table
                max_rows = 200
                current_results = display_df.head(max_rows)
                
                # Dynamic header based on selected languages
                headers_html = "".join([f"<th>{col}</th>" for col in available_cols])
                
                table_html = f"""
                <div class="dict-container">
                    <table class="dict-table">
                        <thead>
                            <tr>
                                <th>Tiếng Việt</th>
                                {headers_html}
                            </tr>
                        </thead>
                        <tbody>
                """
                
                for _, row in current_results.iterrows():
                    v_text = highlight_match(str(row["Vietnamese"]), search_query_norm)
                    table_html += f"<tr><td>{v_text}</td>"
                    for col in available_cols:
                        t_text = highlight_match(str(row[col]), search_query_norm)
                        table_html += f"<td>{t_text}</td>"
                    table_html += "</tr>"
                
                table_html += "</tbody></table></div>"
                
                if len(display_df) > max_rows:
                    st.caption(f"💡 Đang hiện {max_rows}/{len(display_df)} kết quả đầu tiên. Gõ thêm từ khóa để thu hẹp.")
                elif len(display_df) == 0:
                    st.warning("Không tìm thấy kết quả phù hợp.")
                
                st.markdown(table_html, unsafe_allow_html=True)
            else:
                st.error("Từ điển không có cột ngôn ngữ đã chọn.")
    except Exception as e:
        st.error(f"Lỗi khi đọc từ điển: {e}")

# Admin & Template Management Tab Indexing
next_tab_idx = 2

if st.session_state.authenticated and st.session_state.username == "admin":
    with tabs[next_tab_idx]:
        # --- Tab 3: Admin (Admin Only) ---
        st.markdown("### 📈 Quản trị")
        
        # 0. Logging Status
        status_color = "#3498db"
        st.markdown(f"""
        <div style="padding: 15px; border-radius: 10px; border-left: 5px solid {status_color}; background-color: rgba(0,0,0,0.05); margin-bottom: 20px;">
            <div style="font-weight: bold; font-size: 1.1em;">☁️ Nhật ký & tài khoản: <span style="color: {status_color};">Lưu trên Google Sheets</span></div>
            <div style="font-size: 0.9em; color: #666; margin-top: 5px;">
                Cả máy Windows và Streamlit Cloud đều ghi nhật ký và đọc tài khoản từ cùng một Google Sheet (tab đầu tiên: nhật ký, tab <code>Users</code>: tài khoản). Khi không kết nối được, nhật ký tạm ghi vào <code>usage_log.csv</code>.
            </div>
        </div>
        """, unsafe_allow_html=True)

        gs_configured = False
        try:
            if "gsheets" in st.secrets:
                gs_configured = True
            elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                gs_configured = True
        except Exception:
            pass
        
        if not gs_configured:
            st.error("⚠️ **Chưa cấu hình Google Sheets!** Hãy khai báo secrets (file .streamlit/secrets.toml trên máy, hoặc mục Secrets trên Streamlit Cloud).")
            # ... setup guide code ...
        elif st.session_state.get("gsheet_error"):
            st.error(f"❌ **Lỗi ghi nhật ký Google Sheets:** {st.session_state.gsheet_error}")
            if st.button("Xóa thông báo lỗi"):
                del st.session_state["gsheet_error"]
                st.rerun()
            with st.expander("ℹ️ Cách cấu hình ghi nhật ký bằng Google Sheets"):
                st.markdown("""
                1. Tạo Google Service Account và tải khóa JSON.
                2. Chia sẻ Google Sheet với email của service account.
                3. Thêm nội dung sau vào mục **Secrets** của Streamlit Cloud:
                ```toml
                [connections.gsheets]
                spreadsheet = "https://docs.google.com/spreadsheets/d/..."
                type = "service_account"
                project_id = "..."
                private_key_id = "..."
                private_key = "..."
                client_email = "..."
                client_id = "..."
                auth_uri = "https://accounts.google.com/o/oauth2/auth"
                token_uri = "https://oauth2.googleapis.com/token"
                auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
                client_x509_cert_url = "..."
                ```
                """)

        # 1. Translation Tools (including Resolved Dictionary)
        st.markdown("#### 🛠️ Công cụ dịch")
        if st.session_state.get('processed_output_excel') is not None:
            # Use the stored file ID (filename) to generate the download name
            file_ref = st.session_state.get('processed_file_id', 'document').replace('.docx', '')
            st.download_button(
                label="📥 Tải từ điển đã giải mã thẻ (.xlsx)",
                data=st.session_state.processed_output_excel,
                file_name=f"Resolved_Dictionary_{file_ref}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="admin_download_resolved_dict",
                use_container_width=True,
                help="Tải từ điển trong đó các thẻ như [v_name], [e_day] đã được thay bằng giá trị thực."
            )
        else:
            st.info("💡 Chưa có từ điển đã giải mã. Nút tải sẽ xuất hiện sau khi bạn xử lý một báo cáo ở tab 'Xử lý'.")
        
        st.divider()

        with st.expander("📈 Nhật ký sử dụng", expanded=True):
            st.markdown("### Hoạt động gần đây")
            logs = get_logs()
            if st.session_state.get("gsheet_read_error"):
                st.warning(f"Không đọc được Google Sheet ({st.session_state.gsheet_read_error}), "
                           "đang hiển thị nhật ký dự phòng trong usage_log.csv.")
            else:
                st.info("🌐 Đang xem nhật ký từ **Google Sheets**.")
            if logs:
                log_df = pd.DataFrame(logs)
                # Convert Timestamp to datetime for filtering
                log_df['Timestamp'] = pd.to_datetime(log_df['Timestamp'])
                
                # Filter Section UI
                st.markdown("#### 🔍 Bộ lọc")
                c1, c2, c3, c4 = st.columns([1.5, 1, 1, 1.5])
                
                with c1:
                    # Date filter
                    min_date = log_df['Timestamp'].min().date()
                    max_date = log_df['Timestamp'].max().date()
                    # Default to show all
                    date_range = st.date_input(
                        "Khoảng ngày",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date,
                        key="log_filter_date"
                    )
                
                with c2:
                    # User filter
                    all_users = sorted(log_df['User'].unique().tolist())
                    selected_users = st.multiselect("Người dùng", options=all_users, key="log_filter_user")
                
                with c3:
                    # Event Type filter
                    all_events = sorted(log_df['Event Type'].unique().tolist())
                    selected_events = st.multiselect("Loại sự kiện", options=all_events, key="log_filter_event")
                
                with c4:
                    # Keyword search
                    search_query = st.text_input("Tìm trong chi tiết", placeholder="Nhập nội dung...", key="log_filter_search")

                # Apply Filters to the DataFrame
                filtered_df = log_df.copy()
                
                # 1. Filter by Date Range
                if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
                    start_date, end_date = date_range
                    filtered_df = filtered_df[
                        (filtered_df['Timestamp'].dt.date >= start_date) & 
                        (filtered_df['Timestamp'].dt.date <= end_date)
                    ]
                
                # 2. Filter by Users
                if selected_users:
                    filtered_df = filtered_df[filtered_df['User'].isin(selected_users)]
                
                # 3. Filter by Event Types
                if selected_events:
                    filtered_df = filtered_df[filtered_df['Event Type'].isin(selected_events)]
                
                # 4. Filter by Search Keyword
                if search_query:
                    filtered_df = filtered_df[filtered_df['Details'].str.contains(search_query, case=False, na=False)]

                # Display Results
                st.dataframe(filtered_df, use_container_width=True)
                
                # Metrics Row
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    st.info(f"📊 Đang hiện **{len(filtered_df)}** / **{len(log_df)}** bản ghi.")
                
                with col_m2:
                    # Export Filtered Logs to Excel
                    output_log = io.BytesIO()
                    with pd.ExcelWriter(output_log, engine='openpyxl') as writer:
                        # Format timestamp back to string for Excel readability if needed
                        export_df = filtered_df.copy()
                        export_df['Timestamp'] = export_df['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
                        export_df.to_excel(writer, index=False, sheet_name='UsageLogs')
                    log_excel_data = output_log.getvalue()
                    
                    st.download_button(
                        label="📥 Tải nhật ký đã lọc (Excel)",
                        data=log_excel_data,
                        file_name=f"Usage_Log_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_usage_logs",
                        use_container_width=True
                    )
            else:
                st.write("Chưa có nhật ký.")


        with st.expander("👤 Quản lý người dùng", expanded=False):
            st.markdown("### Thêm người dùng mới")
            with st.form("add_user_form", clear_on_submit=True):
                new_user = st.text_input("Tên đăng nhập mới")
                new_pass = st.text_input("Mật khẩu mới", type="password")
                new_role = st.selectbox("Vai trò", ["user", "admin"], index=0)
                add_btn = st.form_submit_button("➕ Thêm người dùng")

                if add_btn:
                    if new_user and new_pass:
                        success, msg = save_user(new_user, new_pass, new_role)
                        if success:
                            st.success(msg)
                            log_event(st.session_state.username, "User Management", f"Created user: {new_user} ({new_role})")
                        else:
                            st.error(msg)
                    else:
                        st.warning("Vui lòng nhập cả tên đăng nhập và mật khẩu.")
            
            st.divider()
            st.markdown("### Người dùng hiện có (sửa/xóa)")
            
            current_users = load_users()
            
            for idx, user in enumerate(current_users):
                is_admin_user = (user["username"] == "admin")
                
                with st.container():
                    # Display individual user row with unique keys
                    c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 1, 1, 1, 1])
                    
                    with c1:
                        # Admin username is protected
                        if is_admin_user:
                            edit_name = st.text_input("Tên", value=user["username"], disabled=True, key=f"un_{idx}")
                        else:
                            edit_name = st.text_input("Tên", value=user["username"], key=f"un_{idx}")
                    
                    with c2:
                        # Password is editable for everyone
                        edit_pass = st.text_input("Mật khẩu", value=user["password"], type="password", key=f"pw_{idx}")
                    
                    with c3:
                        # Role is editable for everyone
                        edit_role = st.selectbox("Vai trò", ["admin", "user"], index=0 if user["role"]=="admin" else 1, key=f"rl_{idx}")
                    
                    with c4:
                        if st.button("🚪", key=f"out_{idx}", use_container_width=True,
                                     help=f"Hủy ghi nhớ đăng nhập của {user['username']} trên mọi máy"):
                            session_store.revoke_user(user["username"])
                            st.success(f"Đã hủy ghi nhớ đăng nhập của {user['username']}.")
                            log_event(st.session_state.username, "User Management", f"Signed out remembered browsers: {user['username']}")

                    with c5:
                        if st.button("💾 Lưu", key=f"upd_{idx}", use_container_width=True):
                            success, msg = update_user_data(user["username"], edit_name, edit_pass, edit_role, False)
                            if success:
                                st.success(msg)
                                log_event(st.session_state.username, "User Management", f"Updated user: {edit_name}")
                                st.rerun()
                            else:
                                st.error(msg)
                    
                    with c6:
                        # Admin account cannot be deleted
                        if is_admin_user:
                            st.button("🚫", disabled=True, key=f"del_dis_{idx}", help="Không thể xóa tài khoản admin chính", use_container_width=True)
                        else:
                            if st.button("🗑️", key=f"del_{idx}", use_container_width=True, help=f"Xóa {user['username']}"):
                                success, msg = remove_user(user["username"])
                                if success:
                                    st.success(msg)
                                    log_event(st.session_state.username, "User Management", f"Deleted user: {user['username']}")
                                    st.rerun()
                                else:
                                    st.error(msg)
                    st.markdown("---")

        with st.expander("⚙️ Cài đặt hệ thống", expanded=False):
            st.markdown("### Cài đặt dịch")
            st.number_input(
                "Ngưỡng phân biệt chữ hoa/thường (độ dài)",
                min_value=1,
                max_value=1000,
                value=st.session_state.case_threshold,
                key="case_threshold",
                help="Câu tiếng Việt ngắn hơn giá trị này sẽ được so khớp có phân biệt chữ hoa/thường; dài bằng hoặc hơn thì không phân biệt."
            )

        with st.expander("📊 Quản lý mẫu Excel", expanded=False):
            st.markdown("### Quản lý mẫu Excel")
            st.markdown("Tải về, cập nhật các file mẫu Excel và đồng bộ dữ liệu hệ thống.")
            
            # Helper for Template Row
            def tab_template_management_row(label, xlsx_path, sync_func, file_name):
                st.markdown(f"**{label}**")
                c1, c2 = st.columns([1, 1])
                with c1:
                    if os.path.exists(xlsx_path):
                        with open(xlsx_path, "rb") as f:
                            st.download_button(
                                label=f"📥 Tải {file_name}",
                                data=f,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"tab_dl_{file_name}",
                                use_container_width=True
                            )
                    else:
                        st.warning(f"Thiếu file {file_name}")
                
                with c2:
                    uploaded = st.file_uploader(f"Cập nhật {file_name}", type=["xlsx"], key=f"tab_up_{file_name}", label_visibility="collapsed")
                    if uploaded:
                        try:
                            with open(xlsx_path, "wb") as f:
                                f.write(uploaded.getbuffer())
                            success, msg = sync_func()
                            if success:
                                st.success(f"✅ Đã cập nhật và đồng bộ {file_name}")
                                log_event(st.session_state.username, "Template Mgmt", f"Uploaded and synced {file_name}")
                            else:
                                st.error(f"Lỗi đồng bộ: {msg}")
                        except Exception as e:
                            st.error(f"Lỗi lưu file: {e}")
                st.divider()

            # 1. Dictionary V3
            tab_template_management_row("1. Từ điển dịch (V3)", tl.DICTIONARY_V3_XLSX, tl.sync_dictionary_v3, "Dictionary_v3.xlsx")
            
            # 2. ParaTemplate
            tab_template_management_row("2. Mẫu đoạn văn", tl.PARA_TEMPLATE_XLSX, tl.sync_para_template, "ParaTemplate.xlsx")
            
            # 3. Text Normalization (CleanV)
            tab_template_management_row("3. Chuẩn hóa văn bản (CleanV)", tl.CLEANV_XLSX, tl.sync_clean_v, "CleanV.xlsx")

            # 4. Confirmed company names (shared table, saved by users from the AI import panel)
            try:
                from company_names_store import dictionary_rows, load_company_names
                company_entries = load_company_names(use_cache=False)
                st.markdown(f"**4. Tên công ty đã xác nhận** ({len(company_entries)} tên)")
                st.caption("Người dùng lưu từ mục 📤 Nhập bản dịch từ AI. Tên được tự thay khi xử lý báo cáo và gửi kèm khi dịch bằng AI. "
                           "Tải về theo mẫu Dictionary_v3 (cột Vietnamese, E, Hs, Ht) để gộp vào từ điển khi cần.")
                if company_entries:
                    st.dataframe([{"Tên tiếng Việt": entry["vietnamese"], "Tên quốc tế": entry["international"],
                                   "Người lưu": entry["saved_by"], "Thời gian": entry["saved_at"]}
                                  for entry in company_entries.values()], hide_index=True, use_container_width=True)
                    company_buffer = io.BytesIO()
                    pd.DataFrame(dictionary_rows(use_cache=False)).to_excel(company_buffer, index=False)
                    st.download_button("📥 Tải tên công ty (mẫu Dictionary_v3)", company_buffer.getvalue(), file_name="Company_names_for_Dictionary_v3.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="tab_company_names_download",
                                       use_container_width=True)
            except Exception as company_error:
                st.warning(f"Không đọc được bảng tên công ty đã xác nhận: {company_error}")
            st.divider()

            # Sync All Button
            if st.button("🔄 Đồng bộ toàn bộ hệ thống", key="tab_sync_all", use_container_width=True, help="Cập nhật toàn bộ file JSON của hệ thống từ các mẫu Excel hiện có."):
                with st.spinner("Đang đồng bộ..."):
                    results = tl.sync_all_templates()
                    all_success = True
                    for name, (success, msg) in results.items():
                        if not success:
                            st.error(f"Lỗi {name}: {msg}")
                            all_success = False
                    if all_success:
                        st.success("✅ Đã đồng bộ toàn bộ dữ liệu từ Excel sang JSON!")
                        log_event(st.session_state.username, "Template Mgmt", "Full synchronization triggered")
                        st.rerun()
    next_tab_idx += 1


# Footer
st.caption("Financial Statements Processor | Ứng dụng xử lý và dịch báo cáo tài chính")









