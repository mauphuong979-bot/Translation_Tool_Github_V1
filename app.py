import streamlit as st
import pandas as pd
import io
import json
import os
import translation_lib as tl
from processor import process_financial_report
from usage_logger import log_event, get_logs
from datetime import datetime

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

# App UI (Authenticated Only)
# Load Initial Metadata from Dictionary.xlsx if not already loaded
if 'metadata_loaded' not in st.session_state:
    try:
        init_metadata, init_df = tl.load_excel_dictionary()
        if init_metadata:
            st.session_state.meta_name_lc = init_metadata.get("name_vn", "")
            st.session_state.meta_name_cap = init_metadata.get("name_trans", "")
            st.session_state.meta_year_end = init_metadata.get("year_end", "")
            st.session_state.meta_date = init_metadata.get("report_date", "")
            st.session_state.meta_period_full = init_metadata.get("period_out", "")
            st.session_state.meta_period_short = init_metadata.get("period_in", "")
            st.session_state.metadata_loaded = True
    except Exception as e:
        st.error(f"Error loading initial metadata: {e}")

st.markdown('<div class="app-logo">📄</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Financial Statements Processor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated analysis and formatting of Word-based financial statements</div>', unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.markdown(f"👤 **Logged in as:** {st.session_state.username}")
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()
        
    st.divider()
    st.markdown("### 📚 Dictionary Management")
    
    DICT_PATH = tl.DICTIONARY_FILE
    
    # Download current dictionary
    if os.path.exists(DICT_PATH):
        try:
             # Provide the ACTUAL file from DICT_PATH to preserve metadata/formulas
             with open(DICT_PATH, "rb") as f:
                 excel_data = f.read()
             
             st.download_button(
                 label="📥 Download Dictionary (Excel)",
                 data=excel_data,
                 file_name="Dictionary.xlsx",
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                 use_container_width=True
             )
        except Exception as e:
             st.error(f"Error preparing download: {e}")
    else:
        st.warning("Dictionary file not found on server.")
        
    # Upload new dictionary
    st.markdown("#### Replace Dictionary")
    new_dict = st.file_uploader("Upload .xlsx", type=["xlsx"], key="dict_upload")
    if new_dict:
        if st.button("🔄 Update Dictionary", use_container_width=True):
            try:
                # Save the uploaded file directly to DICT_PATH
                with open(DICT_PATH, "wb") as f:
                    f.write(new_dict.getbuffer())
                
                # Reload metadata into session state after upload
                new_meta, _ = tl.load_excel_dictionary()
                if new_meta:
                    st.session_state.meta_name_lc = new_meta.get("name_vn", "")
                    st.session_state.meta_name_cap = new_meta.get("name_trans", "")
                    st.session_state.meta_year_end = new_meta.get("year_end", "")
                    st.session_state.meta_date = new_meta.get("report_date", "")
                    st.session_state.meta_period_full = new_meta.get("period_out", "")
                    st.session_state.meta_period_short = new_meta.get("period_in", "")
                
                st.success("Dictionary uploaded and synchronized!")
                log_event(st.session_state.username, "Dictionary", f"Updated from {new_dict.name}")
                st.rerun()
            except Exception as e:
                st.error(f"Error updating dictionary: {e}")

    st.divider()
    st.markdown("### ⚙️ Settings")
    st.info("Additional formatting settings will be added here in the future.")

# Main Interface
with st.expander("📝 **Report Metadata**", expanded=True):
    col_metadata_1, col_metadata_2 = st.columns(2)
    
    with col_metadata_1:
        name_vn_lc = st.text_input("Name in Vietnamese (not capitalized)", placeholder="Công ty TNHH...", key="meta_name_lc")
        reporting_date = st.text_input("Reporting date", placeholder="08/04/2026", key="meta_date")
        
        # Load dictionary columns for translation options
        dict_df = tl.load_dictionary()
        lang_options = ["E", "Hs", "Ht"] # Default options
        if dict_df is not None:
             # Just ensures the columns actually exist in the file
             actual_cols = [c for c in lang_options if c in dict_df.columns]
             if actual_cols:
                 lang_options = actual_cols
                 
        target_lang = st.selectbox("Translate into", options=lang_options, key="meta_translate")
        year_end_date = st.text_input("Year-end date", placeholder="31/12/2025", key="meta_year_end")
        
    with col_metadata_2:
        name_vn_cap = st.text_input("Translated Name (not capitalized)", placeholder="越南嘉泰...", key="meta_name_cap")
        period_full = st.text_input("Period (out of table)", placeholder="từ ngày xx tháng xx năm 20xx đến ngày xx tháng xx năm 20xx", key="meta_period_full")
        period_short = st.text_input("Period (in table)", placeholder="Từ xx/xx/20xx đến xx/xx/20xx", key="meta_period_short")

st.divider()
st.info("📂 **Upload Financial Statements**")
uploaded_file = st.file_uploader("Select file (.docx)", type=["docx"], key="report_file")

if uploaded_file:
    # 1. Action Buttons (Formerly at the bottom)
    col_process_1, col_process_2 = st.columns([1, 2])
    with col_process_1:
        process_btn = st.button("🚀 Process Report", use_container_width=True, type="primary")
    
    if process_btn:
        with st.spinner("Processing report..."):
            try:
                # Prepare selection
                target_col = st.session_state.meta_translate

                # 0. Sync Metadata to Dictionary.xlsx before processing (includes formula recalculation)
                meta_to_save = {
                    "name_vn": st.session_state.meta_name_lc,
                    "name_trans": st.session_state.meta_name_cap,
                    "year_end": st.session_state.meta_year_end,
                    "report_date": st.session_state.meta_date,
                    "period_out": st.session_state.meta_period_full,
                    "period_in": st.session_state.meta_period_short
                }
                tl.save_excel_metadata(meta_to_save)
                
                # Reload dictionary into session state to get recalculated formula results
                _, refreshed_df = tl.load_excel_dictionary()
                if refreshed_df is not None:
                    # Sync UI view with latest calculations
                    display_df = refreshed_df[['Vietnamese', target_col]].dropna().copy()
                    st.session_state.current_dict = display_df

                # Prepare translation map from the ENTIRE current dictionary in session state
                translation_map = dict(zip(st.session_state.current_dict['Vietnamese'], st.session_state.current_dict[target_col]))
                
                # Prepare metadata (with normalization/cleaning)
                metadata = {
                    "Name (not capitalized)": tl.clean_text(st.session_state.meta_name_lc),
                    "Reporting date": tl.clean_text(st.session_state.meta_date),
                    "Translate into": st.session_state.meta_translate,
                    "Year-end date": tl.clean_text(st.session_state.meta_year_end),
                    "Translated Name": tl.clean_text(st.session_state.meta_name_cap),
                    "Period (out of table)": tl.clean_text(st.session_state.meta_period_full),
                    "Period (in table)": tl.clean_text(st.session_state.meta_period_short)
                }
                
                # Call processing logic
                processed_file, msg = process_financial_report(uploaded_file, metadata=metadata, translation_map=translation_map)
                
                if processed_file:
                    st.success(msg)
                    log_event(st.session_state.username, "Processing", f"File: {uploaded_file.name} (Lang: {target_col})")
                    
                    st.divider()
                    st.subheader("Results")
                    st.download_button(
                        label="📥 Download Processed Report (.docx)",
                        data=processed_file,
                        file_name=f"Processed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error(msg)
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")

    st.divider()

    # 2. Dictionary Review & Search
    st.markdown("### 🔍 Dictionary Review & Edit")
    st.info("Review and modify the translations before applying them to the document.")
    
    # Load dictionary if needed
    if 'current_dict' not in st.session_state or st.session_state.get('last_uploaded') != uploaded_file.name:
        dict_df = tl.load_dictionary()
        if dict_df is not None:
            target_col = st.session_state.meta_translate
            if target_col in dict_df.columns:
                display_df = dict_df[['Vietnamese', target_col]].dropna().copy()
                st.session_state.current_dict = display_df
                st.session_state.last_uploaded = uploaded_file.name
            else:
                st.warning(f"Column '{target_col}' not found in dictionary.")
                st.session_state.current_dict = pd.DataFrame(columns=['Vietnamese', target_col])
        else:
            st.error("Could not load Dictionary (Dictionary.xlsx).")
            # Fallback column name
            st.session_state.current_dict = pd.DataFrame(columns=['Vietnamese', st.session_state.meta_translate])

    # Search Logic
    search_query = st.text_input("🔍 Search Keyword (in Vietnamese or Translation)", placeholder="Search...", key="dict_search")
    
    target_col = st.session_state.meta_translate
    
    if search_query:
        # Filter rows that contain the search query in either column
        filtered_df = st.session_state.current_dict[
            st.session_state.current_dict['Vietnamese'].str.contains(search_query, case=False, na=False) |
            st.session_state.current_dict[target_col].str.contains(search_query, case=False, na=False)
        ]
    else:
        filtered_df = st.session_state.current_dict

    # Data Editor
    edited_view = st.data_editor(
        filtered_df,
        num_rows="dynamic",
        use_container_width=True,
        key="dict_editor",
        hide_index=False # Kept index visible to ensure consistent updates if needed
    )
    
    # Update the master dictionary in session state with the edited rows
    if search_query:
        st.session_state.current_dict.update(edited_view)
    else:
        st.session_state.current_dict = edited_view

    # Add Save Button for Dictionary Edits
    if st.button("💾 Save Dictionary Changes to Excel", use_container_width=True):
        try:
            # Load full dictionary to merge changes
            _, full_df = tl.load_excel_dictionary()
            target_col = st.session_state.meta_translate
            
            # Merge edited view back into full_df
            # Note: This simple merge assumes 'Vietnamese' is unique
            for idx, row in st.session_state.current_dict.iterrows():
                full_df.loc[full_df['Vietnamese'] == row['Vietnamese'], target_col] = row[target_col]
            
            # Save metadata and merged df
            meta_to_save = {
                "name_vn": st.session_state.meta_name_lc,
                "name_trans": st.session_state.meta_name_cap,
                "year_end": st.session_state.meta_year_end,
                "report_date": st.session_state.meta_date,
                "period_out": st.session_state.meta_period_full,
                "period_in": st.session_state.meta_period_short
            }
            tl.save_excel_metadata(meta_to_save, full_df)
            st.success("Changes saved and formulas recalculated in Dictionary.xlsx!")
            log_event(st.session_state.username, "Dictionary", "Saved manual edits and recalculated formulas")
            
            # Force reload to show updated formula results
            if 'current_dict' in st.session_state:
                del st.session_state.current_dict
            st.rerun()
        except Exception as e:
            st.error(f"Error saving changes: {e}")
else:
    st.info("Please upload a Word document to start processing.")

# Admin Section: Usage Logs
if st.session_state.authenticated and st.session_state.username == "admin":
    st.divider()
    with st.expander("📈 Usage Logs (Admin Only)", expanded=False):
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
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.info(f"Total actions recorded: {len(logs)}")
        else:
            st.write("No logs found yet.")

    with st.expander("👤 User Management (Admin Only)", expanded=False):
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

# Footer
st.divider()
st.caption("Financial Statements Processor | Professional Streamlit Application")
