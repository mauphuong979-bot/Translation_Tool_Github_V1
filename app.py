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
from usage_logger import log_event, get_logs
from datetime import datetime
import metadata_extractor as mex

# Page Configuration
st.set_page_config(
    page_title="Financial Statements Processor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
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
# Load Users with absolute path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r") as f:
            return json.load(f)["users"]
    return []

def check_credentials(username, password):
    users = load_users()
    for user in users:
        if user["username"] == username and user["password"] == password:
            return True
    return False

def save_user(username, password, role, auto_fill=False):
    users = load_users()
    # Check if user already exists
    if any(u["username"] == username for u in users):
        return False, f"Username '{username}' already exists."
    
    users.append({
        "username": username,
        "password": password,
        "role": role,
        "auto_fill": auto_fill
    })
    
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{username}' added successfully."
    except Exception as e:
        return False, f"Error saving user: {e}"

def remove_user(username):
    users = load_users()
    updated_users = [u for u in users if u["username"] != username]
    
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": updated_users}, f, indent=4)
        return True, f"User '{username}' deleted successfully."
    except Exception as e:
        return False, f"Error deleting user: {e}"

def update_user_data(old_username, new_username, new_password, new_role, auto_fill):
    users = load_users()
    
    # Check if new username is already taken (if changing name)
    if old_username != new_username:
        if any(u["username"] == new_username for u in users):
            return False, f"Username '{new_username}' already exists."
    
    for user in users:
        if user["username"] == old_username:
            user["username"] = new_username
            user["password"] = new_password
            user["role"] = new_role
            user["auto_fill"] = auto_fill
            break
            
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{new_username}' updated successfully."
    except Exception as e:
        return False, f"Error updating user: {e}"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

def handle_autofill():
    """Callback to update password field when user selection changes."""
    users = load_users()
    selected_user = next((u for u in users if u["username"] == st.session_state.login_user), None)
    if selected_user and selected_user.get("auto_fill"):
        st.session_state.login_password = selected_user["password"]
    else:
        st.session_state.login_password = ""

def login_screen():
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-header">🔐 User Login</div>', unsafe_allow_html=True)
    
    users = load_users()
    usernames = [u["username"] for u in users]
    
    # Initialize session state for login widgets if not present
    if "login_user" not in st.session_state:
        st.session_state.login_user = "user" if "user" in usernames else (usernames[0] if usernames else "")
        handle_autofill() # Set initial password if needed

    # User Selection with callback
    st.selectbox("Select Username", usernames, key="login_user", on_change=handle_autofill)
    
    with st.form("login_form"):
        # Password field linked to session state key
        password = st.text_input("Password", key="login_password", type="password")
        submit = st.form_submit_button("Sign In")
        
        if submit:
            if check_credentials(st.session_state.login_user, password):
                st.session_state.authenticated = True
                st.session_state.username = st.session_state.login_user
                log_event(st.session_state.username, "Login", "Successfully signed in")
                st.rerun()
            else:
                st.error("Invalid username or password")
    
    st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# Initialize Settings
if 'processed_file_id' not in st.session_state:
    st.session_state.processed_file_id = None
if 'processed_output_word' not in st.session_state:
    st.session_state.processed_output_word = None
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
    "meta_translate": "E"
}

for key, default_val in metadata_defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default_val



st.markdown('<div class="app-logo">📄</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Financial Statements Processor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated analysis and formatting of Word-based financial statements</div>', unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    # 1. User Profile Section
    st.markdown(f"""
    <div class="sidebar-section">
        <span class="sidebar-label">User Profile</span>
        <div class="sidebar-user-card">
            <div class="sidebar-user-icon">👤</div>
            <div class="sidebar-user-info">
                <div class="sidebar-user-name">{st.session_state.username}</div>
                <div class="sidebar-user-role">{"Administrator" if st.session_state.username == "admin" else "Standard User"}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Sign Out", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 2. Processing Steps Section
    with st.container():
        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown('<span class="sidebar-label">Processing Pipeline</span>', unsafe_allow_html=True)
        st.markdown("### 🛠️ Configuration")
        
        # Initialize process settings in session state if not existing
        if 'process_steps' not in st.session_state:
            st.session_state.process_steps = {
                "unicode": True,
                "clean_v": True,
                "para_template": True,
                "dictionary": True,
                "dual_font": True,
                "number_swap": True,
                "table_size": True,
                "date_format": True,
                "textbox": True,
                "highlight": True
            }

        steps = [
            ("unicode", "1. Unicode Sync (NFC)"),
            ("clean_v", "2. CleanV Fixes"),
            ("para_template", "3. Paragraph Templates"),
            ("dictionary", "4. Dictionary Translation"),
            ("dual_font", "5. Dual-font (CN)"),
            ("number_swap", "6. Number Swap"),
            ("table_size", "7. Table & Layout"),
            ("date_format", "8. Date Formatting"),
            ("textbox", "9. Textbox & Draft"),
            ("highlight", "10. Highlight VN")
        ]
        
        process_settings = {}
        for key, label in steps:
            process_settings[key] = st.checkbox(label, value=st.session_state.process_steps.get(key, True), key=f"step_{key}")
            st.session_state.process_steps[key] = process_settings[key]
        
        st.markdown('</div>', unsafe_allow_html=True)

    # 3. Application Settings Section (at the bottom)
    st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.markdown('<span class="sidebar-label">Preferences</span>', unsafe_allow_html=True)
    st.checkbox(
        "Auto-open/download result", 
        value=st.session_state.get("auto_open", True), 
        key="auto_open", 
        help="Automatically open in Word (local Windows) or trigger download (cloud/Linux) after processing completes."
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Main Interface Tabs Setup
tab_titles = ["🚀 Process", "📖 Dictionary Search"]
if st.session_state.authenticated and st.session_state.username == "admin":
    tab_titles.append("📈 Admin")

tabs = st.tabs(tab_titles)

with tabs[0]:
    # --- Tab 1: Process ---
    st.info("📂 **Upload Financial Statements**")
    uploaded_file = st.file_uploader("Select file (.docx)", type=["docx"], key="report_file")

    # Automatic Metadata Extraction
    if uploaded_file:
        # Use a session state marker to ensure extraction happens only once per new file
        if st.session_state.get('last_extracted_id') != uploaded_file.name:
            # Reset processed data for the new file
            st.session_state.processed_file_id = None
            st.session_state.processed_output_word = None
            st.session_state.processed_output_excel = None
            st.session_state.process_success_msg = None
            
            with st.spinner("Extracting metadata from document..."):
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
                    
                st.session_state.last_extracted_id = uploaded_file.name
                st.rerun() # Refresh to show extracted values in inputs

    st.divider()

    with st.expander("📝 **Report Metadata**", expanded=True):
        col_metadata_1, col_metadata_2 = st.columns(2)
        
        with col_metadata_1:
            # 1. Translate into
            lang_options = ["E", "Hs", "Ht"] 
            default_index = 0
            if st.session_state.meta_translate in lang_options:
                default_index = lang_options.index(st.session_state.meta_translate)
            
            st.selectbox(
                "Translate into", 
                options=lang_options, 
                key="meta_translate",
                index=default_index
            )

            # 2. Translated Name (not capitalized)
            st.text_input("Translated Name (not capitalized)", placeholder="CustomerName", key="meta_name_cap")

            # 3. Name in Vietnamese (not capitalized)
            st.text_input("Name in Vietnamese (not capitalized)", placeholder="ABC Company Co., Ltd...", key="meta_name_lc")
            
        with col_metadata_2:
            # 4. Year-end date
            st.text_input("Year-end date", placeholder="31/12/2025", key="meta_year_end")
            
            # 5. Reporting date
            st.text_input("Reporting date", placeholder="08/04/2026", key="meta_date")

            # 6. Period 1 (in table)
            st.text_input("Period 1 (in table)", placeholder="From xx/xx/20xx to xx/xx/20xx", key="meta_period_short")
            
            # 7. Period 2 (in table)
            st.text_input("Period 2 (in table)", placeholder="From xx/xx/20xx to xx/xx/20xx", key="meta_period_short_2")

    st.divider()

    if uploaded_file:
        st.info(f"📂 **Current file:** {uploaded_file.name}")
        
        # 1. Action Buttons
        col_process_1, col_process_2 = st.columns([1, 2])
        with col_process_1:
            process_btn = st.button("🚀 Process Report", use_container_width=True, type="primary")
        
        # Manual button click
        trigger_processing = process_btn

        if trigger_processing:
            with st.spinner("Processing report..."):
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
                        "period_in_2": st.session_state.meta_period_short_2
                    }
                    
                    v3_df = tl.load_and_fill_v3_dictionary(metadata_for_tags)
                    if v3_df is not None:
                        # Filter and prepare translation map for the specific target language
                        if target_col in v3_df.columns:
                            display_df = v3_df[['Vietnamese', target_col]].dropna().copy()
                            st.session_state.current_dict = display_df
                        else:
                            st.error(f"Target column '{target_col}' not found in Dictionary_v3.")
                            st.stop()
                    else:
                        st.error("Could not load or process Dictionary_v3.xlsx.")
                        st.stop()

                    # Translation map is now built from the dynamically filled dictionary
                    translation_map = dict(zip(st.session_state.current_dict['Vietnamese'], st.session_state.current_dict[target_col]))
                    
                    # Prepare metadata
                    metadata = {
                        "Name (not capitalized)": tl.clean_text(st.session_state.meta_name_lc),
                        "Reporting date": tl.clean_text(st.session_state.meta_date),
                        "Translate into": st.session_state.meta_translate,
                        "Year-end date": tl.clean_text(st.session_state.meta_year_end),
                        "Translated Name": tl.clean_text(st.session_state.meta_name_cap),
                        "Period (in table)": tl.clean_text(st.session_state.meta_period_short)
                    }
                    
                    processed_file, msg = process_financial_report(
                        uploaded_file, 
                        metadata=metadata, 
                        translation_map=translation_map,
                        case_threshold=st.session_state.case_threshold,
                        target_col=target_col,
                        process_settings=st.session_state.process_steps
                    )
                    
                    if processed_file:
                        # Prepare the resolved dictionary for download
                        resolved_dict_output = io.BytesIO()
                        with pd.ExcelWriter(resolved_dict_output, engine='openpyxl') as writer:
                            v3_df.to_excel(writer, index=False, sheet_name='ResolvedDictionary')
                        resolved_dict_data = resolved_dict_output.getvalue()

                        # Generate Dynamic Filename
                        original_name = uploaded_file.name.rsplit('.', 1)[0]
                        target_lang = st.session_state.meta_translate
                        timestamp = datetime.now().strftime("%d%m%y %H%M")
                        processed_filename = f"{original_name}_{target_lang}_tool_{timestamp}.docx"
                        
                        # Save to session state for persistent display
                        st.session_state.processed_file_id = uploaded_file.name
                        st.session_state.processed_output_word = processed_file
                        st.session_state.processed_output_excel = resolved_dict_data
                        st.session_state.process_success_msg = msg
                        st.session_state.processed_filename = processed_filename
                        
                        log_event(st.session_state.username, "Processing", f"File: {uploaded_file.name} (Lang: {target_col})")
                        
                        # AUTO OPEN / DOWNLOAD IF ENABLED
                        if st.session_state.auto_open:
                            if hasattr(os, 'startfile'):
                                # Local Windows Usage: Open Word directly
                                temp_dir = os.path.join(BASE_DIR, "temp_output")
                                if not os.path.exists(temp_dir):
                                    os.makedirs(temp_dir)
                                temp_path = os.path.join(temp_dir, processed_filename)
                                try:
                                    with open(temp_path, "wb") as f:
                                        # Handle both BytesIO and raw bytes
                                        f.write(processed_file.getvalue() if hasattr(processed_file, "getvalue") else processed_file)
                                    os.startfile(temp_path)
                                except Exception as e:
                                    st.warning(f"Could not automatically open Word: {e}. You can still download it manually below.")
                            else:
                                # Cloud / Linux Usage: Trigger automatic browser download
                                st_auto_download(processed_file, processed_filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    else:
                        st.error(msg)
                except Exception as e:
                    st.error(f"An error occurred during processing: {str(e)}")

        # PERSISTENT DOWNLOAD BUTTONS - Shown if processing was successful for current file
        if st.session_state.get('processed_file_id') == uploaded_file.name:
            # Note: Specific success message removed as requested.
            
            st.download_button(
                label="📥 Download Report (.docx)",
                data=st.session_state.processed_output_word,
                file_name=st.session_state.processed_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="btn_download_final_persist",
                use_container_width=True
            )

    st.divider()

with tabs[1]:
    # --- Tab 2: Dictionary Search ---
    st.markdown("### 📖 **Dictionary Search**")
    search_query = st.text_input("🔍 Search terms (Vietnamese or Translation)", placeholder="Enter keyword...", key="dict_search")

    # 1. Load and resolve dictionary with current UI metadata
    metadata_for_display = {
        "name_vn": st.session_state.meta_name_lc,
        "name_trans": st.session_state.meta_name_cap,
        "year_end": st.session_state.meta_year_end,
        "report_date": st.session_state.meta_date,
        "period_in": st.session_state.meta_period_short,
        "period_in_2": st.session_state.meta_period_short_2
    }

    # 2. Language Selection UI
    current_lang = st.session_state.get("meta_translate", "E")
    dict_lang_choices = [
        f"Current ({current_lang})",
        "E",
        "Hs",
        "Ht",
        "All"
    ]
    
    selected_lang_choice = st.radio(
        "Display translation in:",
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
            else: # All (selected_lang_choice == "All")
                target_cols = ["E", "Hs", "Ht"]

            # Filter columns that actually exist in the DataFrame
            available_cols = [c for c in target_cols if c in full_v3_df.columns]
            
            if available_cols:
                cols_to_show = ["Vietnamese"] + available_cols
                display_df = full_v3_df[cols_to_show].copy()
                
                # 4. Filter results based on search query
                if search_query:
                    # Search in Vietnamese and all selected translation columns
                    mask = display_df["Vietnamese"].str.contains(search_query, case=False, na=False)
                    for col in available_cols:
                        mask |= display_df[col].str.contains(search_query, case=False, na=False)
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
                                <th>Vietnamese</th>
                                {headers_html}
                            </tr>
                        </thead>
                        <tbody>
                """
                
                for _, row in current_results.iterrows():
                    v_text = highlight_match(str(row["Vietnamese"]), search_query)
                    table_html += f"<tr><td>{v_text}</td>"
                    for col in available_cols:
                        t_text = highlight_match(str(row[col]), search_query)
                        table_html += f"<td>{t_text}</td>"
                    table_html += "</tr>"
                
                table_html += "</tbody></table></div>"
                
                if len(display_df) > max_rows:
                    st.caption(f"💡 Showing top {max_rows}/{len(display_df)} results. Refine search for more.")
                elif len(display_df) == 0:
                    st.warning("No matching results found.")
                
                st.markdown(table_html, unsafe_allow_html=True)
            else:
                st.error(f"Selected language columns not found in dictionary.")
    except Exception as e:
        st.error(f"Error loading dictionary: {e}")

# Admin & Template Management Tab Indexing
next_tab_idx = 2

if st.session_state.authenticated and st.session_state.username == "admin":
    with tabs[next_tab_idx]:
        # --- Tab 3: Admin (Admin Only) ---
        st.markdown("### 📈 Admin Section")
        
        # 1. Translation Tools (including Resolved Dictionary)
        st.markdown("#### 🛠️ Translation Tools")
        if st.session_state.get('processed_output_excel') is not None:
            # Use the stored file ID (filename) to generate the download name
            file_ref = st.session_state.get('processed_file_id', 'document').replace('.docx', '')
            st.download_button(
                label="📥 Download Resolved Dictionary (.xlsx)",
                data=st.session_state.processed_output_excel,
                file_name=f"Resolved_Dictionary_{file_ref}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="admin_download_resolved_dict",
                use_container_width=True,
                help="Download the dictionary where tags like [v_name], [e_day] are replaced by actual values."
            )
        else:
            st.info("💡 No resolved dictionary available. This button will appear after you process a document in the 'Process' tab.")
        
        st.divider()

        with st.expander("📈 Usage Logs", expanded=True):
            st.markdown("### Recent Tool Activity")
            logs = get_logs()
            if logs:
                log_df = pd.DataFrame(logs)
                st.dataframe(log_df, use_container_width=True)
                
                # Export Logs to Excel
                output_log = io.BytesIO()
                with pd.ExcelWriter(output_log, engine='openpyxl') as writer:
                    log_df.to_excel(writer, index=False, sheet_name='UsageLogs')
                log_excel_data = output_log.getvalue()
                
                st.download_button(
                    label="📥 Download Usage Logs (Excel)",
                    data=log_excel_data,
                    file_name=f"Usage_Log_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_usage_logs"
                )
                
                st.info(f"Total actions recorded: {len(logs)}")
            else:
                st.write("No logs found yet.")

        with st.expander("👤 User Management", expanded=False):
            st.markdown("### Add New User")
            with st.form("add_user_form", clear_on_submit=True):
                new_user = st.text_input("New Username")
                new_pass = st.text_input("New Password", type="password")
                new_role = st.selectbox("Role", ["user", "admin"], index=0)
                new_auto_fill = st.checkbox("Auto-fill Password", value=False)
                add_btn = st.form_submit_button("➕ Add User")
                
                if add_btn:
                    if new_user and new_pass:
                        success, msg = save_user(new_user, new_pass, new_role, new_auto_fill)
                        if success:
                            st.success(msg)
                            log_event(st.session_state.username, "User Management", f"Created user: {new_user} ({new_role})")
                        else:
                            st.error(msg)
                    else:
                        st.warning("Please fill in both username and password.")
            
            st.divider()
            st.markdown("### Current Users (Edit/Delete)")
            
            current_users = load_users()
            
            for idx, user in enumerate(current_users):
                is_admin_user = (user["username"] == "admin")
                
                with st.container():
                    # Display individual user row with unique keys
                    c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 1, 1, 1, 1])
                    
                    with c1:
                        # Admin username is protected
                        if is_admin_user:
                            edit_name = st.text_input("Name", value=user["username"], disabled=True, key=f"un_{idx}")
                        else:
                            edit_name = st.text_input("Name", value=user["username"], key=f"un_{idx}")
                    
                    with c2:
                        # Password is editable for everyone
                        edit_pass = st.text_input("Password", value=user["password"], type="password", key=f"pw_{idx}")
                    
                    with c3:
                        # Role is editable for everyone
                        edit_role = st.selectbox("Role", ["admin", "user"], index=0 if user["role"]=="admin" else 1, key=f"rl_{idx}")
                    
                    with c4:
                        # Auto-fill toggle
                        edit_auto = st.checkbox("Auto-fill", value=user.get("auto_fill", False), key=f"af_{idx}")

                    with c5:
                        if st.button("💾 Save", key=f"upd_{idx}", use_container_width=True):
                            success, msg = update_user_data(user["username"], edit_name, edit_pass, edit_role, edit_auto)
                            if success:
                                st.success(msg)
                                log_event(st.session_state.username, "User Management", f"Updated user: {edit_name}")
                                st.rerun()
                            else:
                                st.error(msg)
                    
                    with c6:
                        # Admin account cannot be deleted
                        if is_admin_user:
                            st.button("🚫", disabled=True, key=f"del_dis_{idx}", help="Cannot delete main admin", use_container_width=True)
                        else:
                            if st.button("🗑️", key=f"del_{idx}", use_container_width=True, help=f"Delete {user['username']}"):
                                success, msg = remove_user(user["username"])
                                if success:
                                    st.success(msg)
                                    log_event(st.session_state.username, "User Management", f"Deleted user: {user['username']}")
                                    st.rerun()
                                else:
                                    st.error(msg)
                    st.markdown("---")

        with st.expander("⚙️ System Settings", expanded=False):
            st.markdown("### Translation Settings")
            st.number_input(
                "Case-sensitivity threshold (length)",
                min_value=1,
                max_value=1000,
                value=st.session_state.case_threshold,
                key="case_threshold",
                help="If the Vietnamese text length is less than this value, translation will be case-sensitive. If greater or equal, it will be case-insensitive."
            )

        with st.expander("📊 Template Management", expanded=False):
            st.markdown("### Excel Template Management")
            st.markdown("Manage Excel templates and synchronize system data.")
            
            # Helper for Template Row
            def tab_template_management_row(label, xlsx_path, sync_func, file_name):
                st.markdown(f"**{label}**")
                c1, c2 = st.columns([1, 1])
                with c1:
                    if os.path.exists(xlsx_path):
                        with open(xlsx_path, "rb") as f:
                            st.download_button(
                                label=f"📥 Download {file_name}",
                                data=f,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"tab_dl_{file_name}",
                                use_container_width=True
                            )
                    else:
                        st.warning(f"Missing {file_name}")
                
                with c2:
                    uploaded = st.file_uploader(f"Update {file_name}", type=["xlsx"], key=f"tab_up_{file_name}", label_visibility="collapsed")
                    if uploaded:
                        try:
                            with open(xlsx_path, "wb") as f:
                                f.write(uploaded.getbuffer())
                            success, msg = sync_func()
                            if success:
                                st.success(f"✅ Updated & synced {file_name}")
                                log_event(st.session_state.username, "Template Mgmt", f"Uploaded and synced {file_name}")
                            else:
                                st.error(f"Sync error: {msg}")
                        except Exception as e:
                            st.error(f"File save error: {e}")
                st.divider()

            # 1. Dictionary V3
            tab_template_management_row("1. Translation Dictionary (V3)", tl.DICTIONARY_V3_XLSX, tl.sync_dictionary_v3, "Dictionary_v3.xlsx")
            
            # 2. ParaTemplate
            tab_template_management_row("2. Paragraph Templates", tl.PARA_TEMPLATE_XLSX, tl.sync_para_template, "ParaTemplate.xlsx")
            
            # 3. Text Normalization (CleanV)
            tab_template_management_row("3. Text Normalization (CleanV)", tl.CLEANV_XLSX, tl.sync_clean_v, "CleanV.xlsx")

            # Sync All Button
            if st.button("🔄 Full System Sync", key="tab_sync_all", use_container_width=True, help="Update all system JSON files from available Excel templates."):
                with st.spinner("Syncing..."):
                    results = tl.sync_all_templates()
                    all_success = True
                    for name, (success, msg) in results.items():
                        if not success:
                            st.error(f"Lỗi {name}: {msg}")
                            all_success = False
                    if all_success:
                        st.success("✅ All data synced from Excel to JSON!")
                        log_event(st.session_state.username, "Template Mgmt", "Full synchronization triggered")
                        st.rerun()
    next_tab_idx += 1


# Footer
st.divider()
st.caption("Financial Statements Processor | Professional Streamlit Application")
