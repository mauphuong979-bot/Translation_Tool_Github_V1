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
        # We provide it as Excel for user convenience when downloading
        try:
             df_current = tl.load_dictionary()
             output_excel = io.BytesIO()
             with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                 df_current.to_excel(writer, index=False)
             excel_data = output_excel.getvalue()
             
             st.download_button(
                 label="📥 Download Dictionary (Excel)",
                 data=excel_data,
                 file_name="General_Dictionary.xlsx",
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                 use_container_width=True
             )
             
             # Also provide raw JSON for developers/GitHub sync
             with open(DICT_PATH, "rb") as f:
                 st.download_button(
                     label="📄 Download Dictionary (JSON)",
                     data=f,
                     file_name="dictionary.json",
                     mime="application/json",
                     use_container_width=True
                 )
        except Exception as e:
             st.error(f"Error preparing download: {e}")
    else:
        st.warning("Dictionary file not found on server.")
        
    # Upload new dictionary
    st.markdown("#### Replace Dictionary")
    new_dict = st.file_uploader("Upload .xlsx or .json", type=["xlsx", "json"], key="dict_upload")
    if new_dict:
        if st.button("🔄 Update Dictionary", use_container_width=True):
            try:
                if new_dict.name.endswith('.xlsx'):
                    df_new = pd.read_excel(new_dict)
                    tl.save_dictionary(df_new)
                else:
                    # Save JSON directly
                    import json
                    content = json.load(new_dict)
                    with open(DICT_PATH, 'w', encoding='utf-8') as f:
                        json.dump(content, f, ensure_ascii=False, indent=4)
                
                st.success("Dictionary updated successfully!")
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
    # Load and Prepare Dictionary for Review
    st.markdown("### 🔍 Dictionary Review & Edit")
    st.info("Review and modify the translations before applying them to the document.")
    
    if 'current_dict' not in st.session_state or st.session_state.get('last_uploaded') != uploaded_file.name:
        dict_df = tl.load_dictionary()
        if dict_df is not None:
            # We only show the Vietnamese and the target language column
            target_col = st.session_state.meta_translate
            if target_col in dict_df.columns:
                display_df = dict_df[['Vietnamese', target_col]].dropna().copy()
                st.session_state.current_dict = display_df
                st.session_state.last_uploaded = uploaded_file.name
            else:
                st.warning(f"Column '{target_col}' not found in dictionary.")
                st.session_state.current_dict = pd.DataFrame(columns=['Vietnamese', 'Translation'])
        else:
            st.error("Could not load Dictionary (dictionary.json).")
            st.session_state.current_dict = pd.DataFrame(columns=['Vietnamese', 'Translation'])

    # Data Editor for the dictionary
    edited_dict = st.data_editor(
        st.session_state.current_dict,
        num_rows="dynamic",
        use_container_width=True,
        key="dict_editor",
        hide_index=True
    )
    
    if st.button("🚀 Process Report"):
        with st.spinner("Processing report..."):
            try:
                # Prepare translation map from edited dictionary
                target_col = st.session_state.meta_translate
                translation_map = dict(zip(edited_dict['Vietnamese'], edited_dict[target_col]))
                
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
                
                # Call processing logic with the edited dictionary map
                processed_file, msg = process_financial_report(uploaded_file, metadata=metadata, translation_map=translation_map)
                
                if processed_file:
                    st.success(msg)
                    
                    # Log event
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
