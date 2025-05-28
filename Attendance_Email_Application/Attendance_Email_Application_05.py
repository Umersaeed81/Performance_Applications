# =============================📦 Import necessary libraries========================
# ============================== 🌐 Streamlit Web App ==============================
import streamlit as st  # 🌐 Streamlit for creating the web interface
# ============================= 📁 File & OS Operations =============================
import os              # 📁 OS operations like path handling
import glob            # 🔍 File searching using patterns
import json            # 📄 JSON file handling
# ============================= ⏱️ Time & Date Utilities =============================
import time                              # ⏱️ Time-related operations (sleep, etc.)
from datetime import datetime, timedelta  # 🕒 Handle dates and time differences
from datetime import datetime             # 🗓️ Datetime for timestamps (redundant import, can remove)
# ============================= 📧 Outlook COM Integration =============================
import pythoncom           # 🧠 Required for working with COM objects on Windows
import win32com.client     # 📧 Access to Outlook via Windows COM interface
# ============================= 📊 Data Processing Libraries =============================
import pandas as pd               # 📊 Data manipulation and analysis
import collections                # 📦 Specialized containers with better performance
from collections import defaultdict  # 🧮 Alternative way to create default dictionaries
# ============================= 🔤 Text & Pattern Handling =============================
import unicodedata  # 🔤 Normalize unicode strings
import re           # 🧹 Regex operations for text cleaning and filtering
# ============================= 💾 Email Settings & Storage =============================
# 📂 Path to the JSON file where filtered email data/settings for Per App 1 will be stored
EMAIL_JSON_per_app1 = "settings_per_app1.json"
# ======================== ⚙️🖥️ Streamlit Page Configuration =========================
# 🔧 Set up the Streamlit page layout, title, and visual theme for Per App 1
def configure_page_per_app1():
    # 🧾 Configure the Streamlit page title and layout
    st.set_page_config(page_title="Outlook Email Downloader", layout="wide")
    
    # 🖼️ Display the main header with maroon color and download icon
    st.markdown(
        "<h1 style='color: maroon;'>📥 Outlook Email Downloader (Email Retrieval)</h1>",
        unsafe_allow_html=True
    )
    
    # 🎨 Insert a horizontal divider with custom style
    st.markdown(
        "<hr style='height:6px; background-color:#8B0000; border:none;'>",
        unsafe_allow_html=True
    )

# ========================== 📧🔌 Initialize Outlook Application ==========================
# 📨 Set up and return the Outlook MAPI interface for accessing email folders (e.g., Inbox)
def initialize_outlook_per_app1():
    # 🧠 Initialize the COM library (required for multi-threaded environments like Streamlit)
    pythoncom.CoInitialize()
    
    # 📬 Launch Outlook and retrieve the MAPI namespace for folder access
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# ============================ 🗂️📂 PST Selection Logic =============================
# 🧠 Handle user selection of the Outlook PST file (current or external)
def select_pst_option_per_app1(outlook_per_app1):
    # ✅ Get PST selection choices from the user
    use_current_pst_per_app1, browse_pst_file_per_app1 = get_user_pst_selection_per_app1()

    # ⚠️ Validate conflicting or missing selections
    if use_current_pst_per_app1 and browse_pst_file_per_app1:
        st.warning("⚠️ Please select only one PST option.")
        return None
    elif not use_current_pst_per_app1 and not browse_pst_file_per_app1:
        st.info("ℹ️ Please select a PST option to proceed.")
        return None

    # 📁 Process the PST file based on user's selection
    if browse_pst_file_per_app1:
        return load_browsed_pst_file_per_app1(outlook_per_app1)
    elif use_current_pst_per_app1:
        return get_current_outlook_inbox_per_app1(outlook_per_app1)

# ========================== 📥 Get PST Selection from User ==========================
# 📥 Subfunction: Capture user checkbox selections for choosing the Outlook PST source
def get_user_pst_selection_per_app1():
    """📝 Display checkboxes to let the user choose between current or browsed PST file"""
    
    # 💼 Option to use the currently loaded Outlook PST
    use_current_per_app1 = st.checkbox("💼 Use current Outlook PST", value=False)
    
    # 🧭 Option to browse and load an external PST file
    browse_pst_per_app1 = st.checkbox("🧭 Browse PST file", value=False)
    
    return use_current_per_app1, browse_pst_per_app1

# ========================== 📂 Load External PST File ==========================
# 📂 Subfunction: Load a user-specified PST file into the Outlook session
def load_browsed_pst_file_per_app1(outlook_per_app1):
    """📂 Load a PST file specified by the user into Outlook and return its root folder"""
    
    # 📝 Prompt the user to enter the full path to the PST file
    pst_path_per_app1 = st.text_input("Enter full path to PST file (e.g., E:\\Email\\LTE_KPI_REPORTING.pst)")
    
    # ✅ Check if a valid path is provided and the file exists
    if pst_path_per_app1 and os.path.exists(pst_path_per_app1):
        try:
            # 📥 Add the PST file to Outlook
            outlook_per_app1.AddStore(pst_path_per_app1)
            
            # 🔍 Search for and return the corresponding root folder
            for i in range(outlook_per_app1.Folders.Count):
                folder_per_app1 = outlook_per_app1.Folders.Item(i + 1)
                if folder_per_app1.FolderPath.endswith(os.path.basename(pst_path_per_app1).replace(".pst", "")):
                    return folder_per_app1
                    
        except Exception as e:
            # ❌ Display error if loading the PST fails
            st.error(f"❌ Error loading PST: {e}")
    
    # ⚠️ Warn user if the path is invalid but not empty
    elif pst_path_per_app1:
        st.warning("⚠️ Invalid PST file path provided.")
    
    return None
# ========================== 📥 Get Current Outlook Inbox ==========================
# 📥 Subfunction: Retrieve the Inbox folder from the currently loaded Outlook profile
def get_current_outlook_inbox_per_app1(outlook_per_app1):
    """📥 Return the Inbox folder from the default Outlook data file (current PST)"""
    
    # 📂 6 = Inbox folder constant in Outlook's MAPI
    return outlook_per_app1.GetDefaultFolder(6)

# ======================= 📁🔍 Retrieve All Subfolders Recursively =======================
# 🔄 Recursively retrieve all subfolders from the given Outlook folder
def get_all_folders_per_app1(folder_per_app1, path_per_app1=""):
    # 🛤️ Build the full folder path (e.g., "Inbox/Subfolder1/Subfolder2")
    full_path_per_app1 = os.path.join(path_per_app1, folder_per_app1.Name) if path_per_app1 else folder_per_app1.Name
    
    # 📦 Start the list with the current folder and its full path
    folders_per_app1 = [(full_path_per_app1, folder_per_app1)]
    
    # 🔁 Recursively traverse and collect subfolders
    for subfolder_per_app1 in folder_per_app1.Folders:
        folders_per_app1.extend(get_all_folders_per_app1(subfolder_per_app1, full_path_per_app1))
    
    # 📤 Return list of tuples (folder_path, folder_object)
    return folders_per_app1

# ============================ 📂✅ Folder Selection UI =============================
# 🖱️ Display checkboxes for each folder so the user can select folders to scan for emails
def get_folder_selections_per_app1(root_folder_per_app1):
    # 📁 Retrieve a list of all folders and subfolders from the selected root folder
    all_folders_per_app1 = get_all_folders_per_app1(root_folder_per_app1)
    
    # 💡 Optional tip to guide users before selecting folders (can be uncommented)
    # st.markdown("> ✅ **Tip:** Select one or more folders to scan for emails.")

    # 📥 Initialize a list to hold user-selected folders
    selected_per_app1 = []

    # 🔘 Generate a checkbox for each folder path and add selected ones to the list
    for folder_path_per_app1, folder_obj_per_app1 in all_folders_per_app1:
        checkbox_selected_per_app1 = st.checkbox(folder_path_per_app1, value=False)
        if checkbox_selected_per_app1:
            # ✅ Add the selected folder (path + object) to the result list
            selected_per_app1.append((folder_path_per_app1, folder_obj_per_app1))
    
    # 📤 Return the final list of selected folders
    return selected_per_app1

# ======================= 📤📧 Retrieve Sender's SMTP Email Address ========================
# 📬 Extract the sender's actual SMTP email address from an Outlook message
def get_sender_smtp_address_per_app1(message_per_app1):
    try:
        # 📧 If the sender's email type is already SMTP, return it directly
        if message_per_app1.SenderEmailType == "SMTP":
            return message_per_app1.SenderEmailAddress or ""
        
        # 🏢 Otherwise, attempt to resolve the Exchange user object
        exch_user_per_app1 = None
        try:
            exch_user_per_app1 = message_per_app1.Sender.GetExchangeUser()
        except Exception as error_per_app1:
            # ❌ Ignore if the sender is not an Exchange user
            pass
        
        # 📤 Try to get the primary SMTP address from the Exchange user object
        if exch_user_per_app1:
            smtp_per_app1 = exch_user_per_app1.PrimarySmtpAddress
            if smtp_per_app1:
                return smtp_per_app1
        
        # 🔁 Fallback to SenderEmailAddress if SMTP could not be resolved
        return message_per_app1.SenderEmailAddress or ""
    
    except Exception as outer_error_per_app1:
        # 🚫 In case of any unexpected failure, return an empty string
        return ""
# =========================== 💾📨 Save Outlook Email as .msg ===========================
# 💌 Save a given Outlook mail item as a .msg file in the specified folder
def save_mail_as_msg_per_app1(mail_per_app1, folder_path_per_app1, filename_per_app1):
    # 🛡️ Clean the filename by removing illegal characters (Windows file system restrictions)
    safe_filename_per_app1 = "".join(
        char_per_app1 for char_per_app1 in filename_per_app1 if char_per_app1 not in r'\\/:*?"<>|'
    )

    # 📁 Construct the full path where the .msg file will be saved
    filepath_per_app1 = os.path.join(folder_path_per_app1, safe_filename_per_app1)
    
    try:
        # 💾 Save the email using format 3 = Outlook .msg file format
        mail_per_app1.SaveAs(filepath_per_app1, 3)
        return True  # ✅ Email saved successfully
    except Exception as error_per_app1:
        return False  # ❌ Failed to save email

# =========================== 💾📝 Save Emails Data to JSON ===========================
# 🗃️ Save the extracted email metadata/content to a JSON file for later use
def save_emails_to_json_per_app1(emails_per_app1):
    # 📂 Open the designated JSON file in write mode (existing content will be overwritten)
    with open(EMAIL_JSON_per_app1, 'w') as f_per_app1:
        # 📝 Serialize the list of email data and write it to the file as JSON
        json.dump(emails_per_app1, f_per_app1)

# ======================= 🧰📧 Email Filtering Options (Sender & Subject) =======================
# 🧪 UI to collect filter options from the user based on sender and subject
def get_email_filters_subject_per_app1():
    # 📧 Section header for sender email filtering
    st.markdown("<h4 style='color: maroon;'>📧 Filter Emails by Sender</h4>", unsafe_allow_html=True)

    # ✅ Checkbox to enable sender email filtering
    check_sender_per_app1 = st.checkbox("Filter by specific sender email?")
    
    # ✉️ Text input field appears only if sender filtering is enabled
    sender_email_per_app1 = st.text_input("Enter sender email") if check_sender_per_app1 else ""

    # 🕵️ Section header for subject filtering
    st.markdown("<h4 style='color: maroon;'>🕵️ Filter Emails by Subject</h4>", unsafe_allow_html=True)

    # ✅ Checkbox to enable subject filtering
    check_subject_per_app1 = st.checkbox("Filter by subject?")
    
    # 🎯 Variables to hold filter mode and criteria value
    subject_filter_mode_per_app1 = None
    subject_value_per_app1 = ""
    
    if check_subject_per_app1:
        # 🧭 Radio options for how the subject should be matched
        subject_filter_mode_per_app1 = st.radio(
            "Choose Subject Filter Type", [
                "Contains keyword",        # 🔍 Subject includes keyword
                "Starts with word",        # 🏁 Subject begins with text
                "Ends with word",          # 🔚 Subject ends with text
                "Exact subject match",     # 🎯 Subject must exactly match
                "Does not contain keyword" # 🚫 Subject must not contain keyword
            ],
            horizontal=True
        )
        # 📝 Input field for subject filtering value
        subject_value_per_app1 = st.text_input("Enter subject criteria:")
    
    # 🔁 Return selected sender and subject filter configurations
    return (
        check_sender_per_app1,
        sender_email_per_app1,
        check_subject_per_app1,
        subject_filter_mode_per_app1,
        subject_value_per_app1
    )

# 📥===============process_emails_per_app1: Main Email Processing Function==================📥
def process_emails_per_app1(
    selected_folders_per_app1, sender_emails_per_app1, date_range_per_app1,
    download_path_per_app1, size_preference_per_app1, apply_size_filter_per_app1,
    size_condition_per_app1, size_value_per_app1, size_range_per_app1,
    download_mode_per_app1="Delete all files if exist.",
    apply_subject_filter_per_app1=False, subject_filter_mode_per_app1=None, subject_value_per_app1="",
    subject_case_sensitive_per_app1=False,
    apply_read_filter_per_app1=False, read_status_per_app1="Unread",
    apply_attachment_filter_per_app1=False, attachment_filter_type_per_app1=None,
    attachment_extensions_per_app1=None,
    attachment_name_filter_mode_per_app1=None, attachment_name_filter_value_per_app1="",
    apply_importance_filter_per_app1=False, selected_importance_per_app1="Normal",
    apply_flag_filter_per_app1=False, flag_status_value_per_app1="No Flag"
):
    if attachment_extensions_per_app1 is None:
        attachment_extensions_per_app1 = []

    # 🔠==============================Text Normalization==============================
    def normalize_text_per_app1(text_per_app1, case_sensitive_per_app1=False):
        """Normalize text with optional case sensitivity."""
        norm_per_app1 = unicodedata.normalize("NFKC", text_per_app1 or "").strip()
        return norm_per_app1 if case_sensitive_per_app1 else norm_per_app1.lower()

    # 🧹==============================Filename Sanitization==============================
    def sanitize_filename_per_app1(filename_per_app1):
        """Replace illegal characters and extra spaces in filenames."""
        filename_per_app1 = re.sub(r'\s+', ' ', filename_per_app1)
        sanitized_per_app1 = re.sub(r'[\\/*?:"<>|]', "_", filename_per_app1)
        return sanitized_per_app1.strip()

    # 🔁==============================Generate Unique Filename==============================
    def get_unique_filename_per_app1(path_per_app1):
        """Return a non-colliding filepath by appending a counter."""
        base_per_app1, ext_per_app1 = os.path.splitext(path_per_app1)
        counter_per_app1 = 1
        new_path_per_app1 = path_per_app1
        while os.path.exists(new_path_per_app1):
            new_path_per_app1 = f"{base_per_app1} ({counter_per_app1}){ext_per_app1}"
            counter_per_app1 += 1
        return new_path_per_app1

    # 📁==============================Prepare Download Directory==============================
    def prepare_download_directory_per_app1(path_per_app1, mode_per_app1):
        """Create or clean the download directory based on user mode."""
        if not os.path.exists(path_per_app1):
            os.makedirs(path_per_app1)
        if mode_per_app1 == "Delete all files if exist.":
            files_per_app1 = glob.glob(os.path.join(path_per_app1, '*'))
            for f_per_app1 in files_per_app1:
                try:
                    if os.path.isfile(f_per_app1):
                        os.remove(f_per_app1)
                except Exception as e_per_app1:
                    st.error(f"❌ Failed to delete {f_per_app1}: {e_per_app1}")

    # 🧠==============================Subject Filter Logic==============================
    def passes_subject_filter_per_app1(mail_per_app1):
        """Apply subject filter with multiple modes."""
        if not apply_subject_filter_per_app1:
            return True
        subject_per_app1 = mail_per_app1.Subject or ""
        norm_subject_per_app1 = normalize_text_per_app1(subject_per_app1, subject_case_sensitive_per_app1)
        keywords_per_app1 = [
            normalize_text_per_app1(val_per_app1, subject_case_sensitive_per_app1)
            for val_per_app1 in subject_value_per_app1.split(",") if val_per_app1.strip()
        ]
        for keyword_per_app1 in keywords_per_app1:
            if subject_filter_mode_per_app1 == "🔍 Contains keyword" and keyword_per_app1 in norm_subject_per_app1:
                return True
            if subject_filter_mode_per_app1 == "🏁 Starts with word" and norm_subject_per_app1.startswith(keyword_per_app1):
                return True
            if subject_filter_mode_per_app1 == "🏁➡️Ends with word" and norm_subject_per_app1.endswith(keyword_per_app1):
                return True
            if subject_filter_mode_per_app1 == "🎯 Exact subject match" and norm_subject_per_app1 == keyword_per_app1:
                return True
            if subject_filter_mode_per_app1 == "🚫🔍 Does not contain keyword" and keyword_per_app1 in norm_subject_per_app1:
                return False
        return subject_filter_mode_per_app1 == "🚫🔍 Does not contain keyword"

    # 📏==============================Size Filter Logic==============================
    def passes_size_filter_per_app1(mail_per_app1):
        """Filter emails by size using various conditions."""
        if not apply_size_filter_per_app1:
            return True
        size_per_app1 = mail_per_app1.Size
        if size_condition_per_app1 == "=" and size_per_app1 != size_value_per_app1:
            return False
        if size_condition_per_app1 == "!=" and size_per_app1 == size_value_per_app1:
            return False
        if size_condition_per_app1 == "<" and size_per_app1 >= size_value_per_app1:
            return False
        if size_condition_per_app1 == ">" and size_per_app1 <= size_value_per_app1:
            return False
        if size_condition_per_app1 == "<=" and size_per_app1 > size_value_per_app1:
            return False
        if size_condition_per_app1 == ">=" and size_per_app1 < size_value_per_app1:
            return False
        if size_condition_per_app1 == "Between" and not (size_range_per_app1[0] <= size_per_app1 <= size_range_per_app1[1]):
            return False
        return True

    # 👀==============================Read Status Filter==============================
    # def passes_read_filter_per_app1(mail_per_app1):
    #     """Filter emails based on read/unread status."""
    #     if not apply_read_filter_per_app1:
    #         return True
    #     return mail_per_app1.UnRead if read_status_per_app1 == "Unread" else not mail_per_app1.UnRead
    def passes_read_filter_per_app1(mail_per_app1):
        """📖 Filter emails based on selected read/unread status."""
        if not apply_read_filter_per_app1:
            return True
    
        if read_status_per_app1 == "📩 Unread":
            return mail_per_app1.UnRead
        elif read_status_per_app1 == "📖 Read":
            return not mail_per_app1.UnRead
    
        return True  # Fallback (shouldn't occur if UI is correct)

    # ❗==============================Importance Filter==============================
    def passes_importance_filter_per_app1(mail_per_app1):
        """Filter emails by importance level."""
        if not apply_importance_filter_per_app1:
            return True
        importance_per_app1 = mail_per_app1.Importance
        if selected_importance_per_app1 == "🟢 Low" and importance_per_app1 != 0:
            return False
        if selected_importance_per_app1 == "🟡 Normal" and importance_per_app1 != 1:
            return False
        if selected_importance_per_app1 == "🔴 High" and importance_per_app1 != 2:
            return False
        return True
    # 🚩==============================Follow-Up Flag Filter==============================
    def passes_flag_filter_per_app1(mail_per_app1):
        """Filter emails by follow-up flag status."""
        if not apply_flag_filter_per_app1:
            return True
        flag_map_per_app1 = {"🚫 No Flag": 0, "✅ Completed": 1, "🚩 Flagged": 2}
        return mail_per_app1.FlagStatus == flag_map_per_app1.get(flag_status_value_per_app1, 0)

    # 📎==============================Attachment Filter==============================
    def passes_attachment_filter_per_app1(mail_per_app1):
        """Filter emails with or without attachments by extension and name."""
        if not apply_attachment_filter_per_app1:
            return True
        has_attachments_per_app1 = bool(mail_per_app1.Attachments.Count > 0)
        if attachment_filter_type_per_app1 == "With Attachments Only":
            if not has_attachments_per_app1:
                return False
            for i_per_app1 in range(1, mail_per_app1.Attachments.Count + 1):
                try:
                    attachment_per_app1 = mail_per_app1.Attachments.Item(i_per_app1)
                    filename_per_app1 = attachment_per_app1.FileName or ""
                    filename_norm_per_app1 = normalize_text_per_app1(filename_per_app1)
                    if attachment_extensions_per_app1:
                        if not any(filename_per_app1.lower().endswith(ext_per_app1.lower()) for ext_per_app1 in attachment_extensions_per_app1):
                            continue
                    if attachment_name_filter_value_per_app1.strip():
                        filter_text_per_app1 = normalize_text_per_app1(attachment_name_filter_value_per_app1)
                        if attachment_name_filter_mode_per_app1 == "🔍 Contains" and filter_text_per_app1 not in filename_norm_per_app1:
                            continue
                        if attachment_name_filter_mode_per_app1 == "🚫🔍 Does not contain" and filter_text_per_app1 in filename_norm_per_app1:
                            continue
                        if attachment_name_filter_mode_per_app1 == "🏁 Starts with" and not filename_norm_per_app1.startswith(filter_text_per_app1):
                            continue
                        if attachment_name_filter_mode_per_app1 == "🏁➡️ Ends with" and not filename_norm_per_app1.endswith(filter_text_per_app1):
                            continue
                        if attachment_name_filter_mode_per_app1 == "🎯 Exact match" and filename_norm_per_app1 != filter_text_per_app1:
                            continue
                    return True
                except Exception:
                    continue
            return False
        elif attachment_filter_type_per_app1 == "Without Attachments Only":
            return not has_attachments_per_app1
        return True

    # 📤============================== Get Filtered Emails==============================
    def get_filtered_emails_per_app1():
        """Retrieve and filter emails from Outlook folders based on all criteria."""
        emails_by_date_sender_per_app1 = collections.defaultdict(list)
        start_dt_per_app1 = datetime.combine(min(date_range_per_app1), datetime.min.time())
        end_dt_per_app1 = datetime.combine(max(date_range_per_app1), datetime.max.time())

        for _, folder_per_app1 in selected_folders_per_app1:
            try:
                restriction_per_app1 = f"[ReceivedTime] >= '{start_dt_per_app1.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_dt_per_app1.strftime('%m/%d/%Y %H:%M %p')}'"
                items_per_app1 = folder_per_app1.Items.Restrict(restriction_per_app1)
            except Exception:
                items_per_app1 = folder_per_app1.Items

            for i_per_app1 in range(items_per_app1.Count):
                try:
                    mail_per_app1 = items_per_app1.Item(i_per_app1 + 1)
                    if mail_per_app1.Class != 43:
                        continue
                    smtp_per_app1 = get_sender_smtp_address_per_app1(mail_per_app1).lower()
                    if smtp_per_app1 not in [s.lower() for s in sender_emails_per_app1]:
                        continue
                    if not passes_subject_filter_per_app1(mail_per_app1): continue
                    if not passes_size_filter_per_app1(mail_per_app1): continue
                    if not passes_read_filter_per_app1(mail_per_app1): continue
                    if not passes_importance_filter_per_app1(mail_per_app1): continue
                    if not passes_flag_filter_per_app1(mail_per_app1): continue
                    if not passes_attachment_filter_per_app1(mail_per_app1): continue
                    received_date_per_app1 = mail_per_app1.ReceivedTime.date()
                    emails_by_date_sender_per_app1[(received_date_per_app1, smtp_per_app1)].append(mail_per_app1)
                except Exception:
                    continue
        return emails_by_date_sender_per_app1

    # 💾==============================Save Email to File==============================
    def save_email_per_app1(mail_per_app1, filename_base_per_app1, index_per_app1=None):
        """Save email as a .msg file with unique or safe filename."""
        subject_per_app1 = mail_per_app1.Subject or "NoSubject"
        sender_per_app1 = mail_per_app1.SenderName or "UnknownSender"
        if index_per_app1 is not None:
            filename_per_app1 = f"{filename_base_per_app1} - {sender_per_app1} - {subject_per_app1} ({index_per_app1+1}).msg"
        else:
            filename_per_app1 = f"{filename_base_per_app1} - {sender_per_app1} - {subject_per_app1}.msg"
        safe_filename_per_app1 = sanitize_filename_per_app1(filename_per_app1)
        filepath_per_app1 = os.path.join(download_path_per_app1, safe_filename_per_app1)
        if download_mode_per_app1 == "Do not overwrite, rename it.":
            filepath_per_app1 = get_unique_filename_per_app1(filepath_per_app1)
        try:
            time.sleep(0.1)
            mail_per_app1.SaveAs(filepath_per_app1, 3)
            return True
        except Exception as e_per_app1:
            fallback_per_app1 = getattr(mail_per_app1, "Subject", "Unknown Subject")
            st.warning(f"⚠️ Could not save: {fallback_per_app1[:40]}")
            st.error(f"❌ Failed to save email: {e_per_app1}")
            return False

    # 📆==============================Process Emails by Date and Sender==============================
    def process_emails_by_day_per_app1(emails_by_date_sender_per_app1):
        """Download filtered emails and build summary of missing ones."""
        summary_per_app1 = []
        counter_per_app1 = 0
        for day_per_app1 in date_range_per_app1:
            for sender_per_app1 in sender_emails_per_app1:
                key_per_app1 = (day_per_app1.date(), sender_per_app1.lower())
                mails_per_app1 = emails_by_date_sender_per_app1.get(key_per_app1, [])
                if mails_per_app1:
                    if apply_size_filter_per_app1:
                        for idx, mail in enumerate(mails_per_app1):
                            if save_email_per_app1(mail, day_per_app1.strftime('%Y-%m-%d'), idx):
                                counter_per_app1 += 1
                    elif size_preference_per_app1 in ["🔽 Smallest", "🔼 Largest"]:
                        mail = min(mails_per_app1, key=lambda m: m.Size) if size_preference_per_app1 == "🔽 Smallest" else max(mails_per_app1, key=lambda m: m.Size)
                        if save_email_per_app1(mail, day_per_app1.strftime('%Y-%m-%d')):
                            counter_per_app1 += 1
                    else:
                        for idx, mail in enumerate(mails_per_app1):
                            if save_email_per_app1(mail, day_per_app1.strftime('%Y-%m-%d'), idx):
                                counter_per_app1 += 1
                else:
                    summary_per_app1.append({
                        "Date": day_per_app1.strftime('%Y-%m-%d'),
                        "Missing Email Address": sender_per_app1
                    })
        return summary_per_app1, counter_per_app1

    # 📊==============================Export Missing Email Summary==============================
    def export_missing_summary_per_app1(summary_per_app1):
        """Save missing email summary to Excel file."""
        path_per_app1 = os.path.join(download_path_per_app1, "Missing_Email_Summary.xlsx")
        df_per_app1 = pd.DataFrame(summary_per_app1)
        df_per_app1.to_excel(path_per_app1, index=False)

    # 🚀==============================Execute the Workflow==============================
    with st.spinner("⏳ Processing emails..."):
        prepare_download_directory_per_app1(download_path_per_app1, download_mode_per_app1)
        filtered_emails_per_app1 = get_filtered_emails_per_app1()
        summary_per_app1, count_per_app1 = process_emails_by_day_per_app1(filtered_emails_per_app1)
        export_missing_summary_per_app1(summary_per_app1)
        return summary_per_app1, count_per_app1

# 📂==============================Load Previously Saved Email Metadata (if any)==============================📂
def load_saved_emails_per_app1():
    """
    Load previously saved email metadata from a JSON file.

    Returns:
        list: A list of saved email metadata, or an empty list if the file doesn't exist.
    """
    # ✅ Check if the saved email metadata file exists
    if os.path.exists(EMAIL_JSON_per_app1):
        # 📖 Open and read the JSON file
        with open(EMAIL_JSON_per_app1, 'r') as f_per_app1:
            return json.load(f_per_app1)
    
    # 🚫 Return empty list if no saved data is found
    return []

# ⚙️📂==============================Configure Download Settings and Folder Selection UI==============================📂⚙️
def get_download_settings_per_app1(default_download_path_per_app1="D:\\Downloaded_Attachments"):
    """
    Display UI for configuring download folder and mode in Streamlit.

    Args:
        default_download_path_per_app1 (str): The default path where attachments will be downloaded.

    Returns:
        tuple: (download_path_per_app1, download_mode_per_app1)
               download_path_per_app1 (str): Selected folder path.
               download_mode_per_app1 (str): Selected download mode option.
    """
    # 📝 Section Header: Download Folder Path
    st.markdown("<h4 style='color: maroon;'>📂 Select Download Location</h4>", unsafe_allow_html=True)
    
    # 🧾 Input field to specify download path (default provided)
    download_path_per_app1 = st.text_input("📂 Download Folder Path:", value=default_download_path_per_app1)
    
    # 🗂️ Create folder if it doesn't exist (no error if it already exists)
    os.makedirs(download_path_per_app1, exist_ok=True)

    # 🛠️ Section Header: Download Mode Options
    st.markdown("<h4 style='color: maroon;'>🛠️ Download Mode Configuration</h4>", unsafe_allow_html=True)
    
    # ⚙️ User selects how to handle existing files in the folder
    download_mode_per_app1 = st.selectbox("Choose download mode ⚙️:", [
        "Delete all files if exist.",           # 🚮 Remove everything before saving
        "Do not overwrite, rename it."          # 📝 Save with new names if conflicts
    ])
    
    # 📤 Return both path and mode for downstream processing
    return download_path_per_app1, download_mode_per_app1

# 📝🔄==============================Generate Unique Filename for Attachments==============================🔄📝
def get_unique_filename_per_app1(file_path_per_app1):
    """
    Generate a unique file path by appending a counter if a file already exists.

    Args:
        file_path_per_app1 (str): Original file path.

    Returns:
        str: A new unique file path that doesn't conflict with existing files.
    """
    # 🧱 Split the original file path into base name and extension
    base_per_app1, ext_per_app1 = os.path.splitext(file_path_per_app1) 
    
    # 🔢 Counter for filename versioning
    counter_per_app1 = 1

    # 📁 Start with the original path
    new_path_per_app1 = file_path_per_app1

    # 🔄 Check if file already exists, if yes, append counter to make it unique
    while os.path.exists(new_path_per_app1):
        # 🆕 Create new versioned filename like file_1.ext, file_2.ext, etc.
        new_path_per_app1 = f"{base_per_app1}_{counter_per_app1}{ext_per_app1}"
        counter_per_app1 += 1

    # ✅ Return a unique file path that doesn't currently exist
    return new_path_per_app1
#============================== 🚀🖥️ Main Application Logic ==============================
# 🗂️🔧=======================1. Configure Folder Selection UI (Left Column)==============================🔧🗂️
def handle_folder_selection_ui_per_app1(outlook_per_app1):
    """
    Display the UI for selecting PST and email folders using Streamlit.

    Args:
        outlook_per_app1: Outlook application object used for navigating PST files.

    Returns:
        Tuple: (root_folder, list of selected folders, confirmation status)
    """
    # 📌 Section Header: Select PST Option
    st.markdown("<h4 style='color: maroon;'>📌 Select PST Option</h4>", unsafe_allow_html=True)
    
    # 📂 Let user pick the PST file/folder root
    root_folder_per_app1 = select_pst_option_per_app1(outlook_per_app1)

    # ✅ Initialize confirmation and folder list
    confirm_selection_per_app1 = False
    selected_folders_per_app1 = []

    # 🗃️ If a PST folder was selected
    if root_folder_per_app1:
        # Section Header: Select Email Folders
        st.markdown("<h4 style='color: maroon;'>🗃️ Select Folder(s) to Scan for Emails</h4>", unsafe_allow_html=True)
        
        # 📥 Show folder tree for user to select from
        selected_folders_per_app1 = get_folder_selections_per_app1(root_folder_per_app1)
        
        if not selected_folders_per_app1:
            # ℹ️ Prompt user to select at least one folder
            st.info("ℹ️ Please select at least one folder before confirming.")
        else:
            # ☑️ User confirms their folder selection
            confirm_selection_per_app1 = st.checkbox("✅ Are you sure you have selected all the required folders?")
            if confirm_selection_per_app1:
                st.success(f"📂 You have selected **{len(selected_folders_per_app1)}** folder(s).")

    return root_folder_per_app1, selected_folders_per_app1, confirm_selection_per_app1

# 🧑‍💼✉️==============================2. Sender & Subject Filter Configuration==============================
def get_sender_and_subject_filters_per_app1(saved_emails_per_app1):
    """
    Displays Streamlit UI for configuring sender email filters and optional subject filters.

    Args:
        saved_emails_per_app1 (list): List of saved sender email addresses to pre-populate input.

    Returns:
        tuple: 
            - new_emails (list): List of entered sender email addresses
            - check_subject (bool): Whether subject filtering is enabled
            - subject_filter_mode (str|None): Chosen filter mode (contains, starts with, etc.)
            - subject_value (str): Subject keyword entered by the user
            - subject_case_sensitive (bool): Case sensitivity setting for subject filtering
    """
    # 📧 Section Header: Sender Filter
    st.markdown("<h4 style='color: maroon;'>🆔 Sender Email Address</h4>", unsafe_allow_html=True)
    
    # 📥 Input field to enter comma-separated sender emails
    email_input_per_app1 = st.text_area(
        "✉️ Enter Sender Email Addresses (comma separated):", 
        ", ".join(saved_emails_per_app1)
    )
    
    # 🧹 Clean and split email addresses
    new_emails_per_app1 = [
        e_per_app1.strip() 
        for e_per_app1 in email_input_per_app1.split(',') 
        if e_per_app1.strip()
    ]

    # 💾 Save button for storing entered emails
    if st.button("💾 Save Email Addresses"):
        save_emails_to_json_per_app1(new_emails_per_app1)
        st.success("✅ Email addresses saved!")

    # 🧠 Section Header: Subject Filter
    st.markdown("<h4 style='color: maroon;'>🕵️ Filter Emails by Subject</h4>", unsafe_allow_html=True)
    
    # ✅ Enable subject filter
    check_subject_per_app1 = st.checkbox("(optional) Filter by subject🔤?")
    
    # Initialize subject filter config
    subject_filter_mode_per_app1 = None
    subject_value_per_app1 = ""
    subject_case_sensitive_per_app1 = False

    if check_subject_per_app1:
        # 🧭 Choose type of subject filter
        subject_filter_mode_per_app1 = st.radio(
            "🔍 Choose Subject Filter Type",
            [
                "🔍 Contains keyword", 
                "🏁 Starts with word", 
                "🏁➡️ Ends with word", 
                "🎯 Exact subject match", 
                "🚫🔍 Does not contain keyword"
            ],
            horizontal=True
        )
        
        # 📝 Input field for subject keyword
        subject_value_per_app1 = st.text_input("📬 Enter subject criteria:")
        
        # 🔠 Case sensitivity option
        subject_case_sensitive_per_app1 = st.checkbox("🔠 Make subject filter case sensitive")

    # 🧾 Return all filter settings
    return (
        new_emails_per_app1, 
        check_subject_per_app1, 
        subject_filter_mode_per_app1, 
        subject_value_per_app1.strip(), 
        subject_case_sensitive_per_app1
    )
# 📅============================== 3. Date Range Selection UI==============================
def get_date_range_per_app1():
    """
    Displays a Streamlit UI for selecting a start and end date, and generates the full date range.

    Returns:
        tuple:
            - start_date (date): Selected start date
            - end_date (date): Selected end date
            - date_list (list[datetime]): List of all dates in the selected range as datetime objects
    """
    # 📆 Section Header
    st.markdown("<h4 style='color: maroon;'>📆 Select Date Range</h4>", unsafe_allow_html=True)

    # 🪟 Create two side-by-side columns for start and end date
    d1_per_app1, d2_per_app1 = st.columns(2)

    with d1_per_app1:
        # 🟢 Start date input
        start_date_per_app1 = st.date_input("🟢📆 Start Date", key="start_date_per_app1")

    with d2_per_app1:
        # 🔴 End date input
        end_date_per_app1 = st.date_input("🔴📆 End Date", key="end_date_per_app1")

    # 📅 Generate a list of datetime objects from start to end date (inclusive)
    date_range_per_app1 = pd.date_range(start=start_date_per_app1, end=end_date_per_app1).to_pydatetime().tolist()

    # 📤 Return values for filtering use
    return start_date_per_app1, end_date_per_app1, date_range_per_app1
# 📏============================== 4. Email Size Filter Options UI==============================
def get_size_filter_options_per_app1():
    """
    Displays Streamlit UI for configuring email size-based filtering.

    Returns:
        tuple:
            - size_selection_mode_per_app1 (str or None)
            - size_preference_per_app1 (str or None)
            - apply_size_filter_per_app1 (bool)
            - size_condition_per_app1 (str or None)
            - size_value_per_app1 (int or None)
            - size_range_per_app1 (tuple[int, int] or None)
    """
    # 📌 Section Header
    st.markdown("<h4 style='color: maroon;'>📏 Email Size Options</h4>", unsafe_allow_html=True)

    # 📦 Checkbox to enable size filter logic
    apply_size_logic_per_app1 = st.checkbox("(optional) Apply Email Size Options? 📏")

    # 🔧 Initialize all return variables
    size_selection_mode_per_app1 = size_preference_per_app1 = size_condition_per_app1 = None
    size_value_per_app1 = size_range_per_app1 = None
    apply_size_filter_per_app1 = False

    # ⚙️ Display size filter options only if enabled
    if apply_size_logic_per_app1:
        # 📊 Choose between saving smallest/largest OR applying custom filter
        size_selection_mode_per_app1 = st.radio(
            "☑️ Select an option:",
            ["📏 Choose email size to save", "⚙️ Enable Email Size Filtering"], horizontal=True,
            index=0
        )

        if size_selection_mode_per_app1 == "📏 Choose email size to save":
            # 🎯 User wants to extract only the smallest or largest email
            size_preference_per_app1 = st.selectbox("📏 Choose email size to save:", ["🔽 Smallest", "🔼 Largest"])

        elif size_selection_mode_per_app1 == "⚙️ Enable Email Size Filtering":
            # 🔍 Enable advanced size filtering
            apply_size_filter_per_app1 = True
            #st.markdown("<h4 style='color: maroon;'>📐 Email Size Filter</h4>", unsafe_allow_html=True)

            # 🧮 Comparison condition selection (UI only)
            condition_options_per_app1 = {
                "🟰 Equal": "=",
                "🚫🟰 Not Equal": "!=",
                "🔽 Less Than": "<",
                "🔼 Greater Than": ">",
                "⬇️🟰 Less Than or Equal": "<=",
                "⬆️🟰 Greater Than or Equal": ">=",
                "↔️ Between": "Between"
            }

            # Show user-friendly labels in the UI
            condition_label_per_app1 = st.selectbox("🧮 Choose comparison condition:", list(condition_options_per_app1.keys()))
            size_condition_per_app1 = condition_options_per_app1[condition_label_per_app1]

            if size_condition_per_app1 == "Between":
                # ➕ Range input for min and max size
                size_min_per_app1 = st.number_input("Min Size (KB)", min_value=0, step=1)
                size_max_per_app1 = st.number_input("Max Size (KB)", min_value=0, step=1)
                size_range_per_app1 = (size_min_per_app1 * 1024, size_max_per_app1 * 1024)  # Convert to bytes
            else:
                # ➖ Single value input
                size_value_per_app1 = st.number_input("📦 Email Size (KB)", min_value=0, step=1) * 1024  # Convert to bytes

    # 📨 Return all size filter parameters
    return size_selection_mode_per_app1, size_preference_per_app1, apply_size_filter_per_app1, size_condition_per_app1, size_value_per_app1, size_range_per_app1


#==============================🧠 Final Refactored main_per_app1() Function==============================
def main_per_app1():
    # 🧱 Set up the Streamlit page configuration and initialize Outlook application
    configure_page_per_app1()
    outlook_per_app1 = initialize_outlook_per_app1()

    # 📊 Create two columns for UI layout
    col1_per_app1, col2_per_app1 = st.columns(2)

    #==============================📁 Folder Selection Panel (Left Column)==============================
    with col1_per_app1:
        root_folder_per_app1, selected_folders_per_app1, confirm_selection_per_app1 = handle_folder_selection_ui_per_app1(outlook_per_app1)

    # ✅ Proceed only if folders are selected and user confirms
    if confirm_selection_per_app1 and selected_folders_per_app1:

        #==============================⚙️ Filter Configuration Panel (Right Column)==============================
        with col2_per_app1:
            
            # 📆 Date Range Selection
            start_date_per_app1, end_date_per_app1, date_range_per_app1 = get_date_range_per_app1()         

            # 🆔 Sender & Subject Filters
            saved_emails_per_app1 = load_saved_emails_per_app1()
            new_emails_per_app1, check_subject_per_app1, subject_filter_mode_per_app1, subject_value_per_app1, subject_case_sensitive_per_app1 = get_sender_and_subject_filters_per_app1(saved_emails_per_app1)

       

            # 📖 Read/Unread Filter (Optional)
            st.markdown("<h4 style='color: maroon;'>📖 Read / Unread Filter</h4>", unsafe_allow_html=True)
            apply_read_filter_per_app1 = st.checkbox("(optional) Apply read/unread filter? 👁️")
            read_status_per_app1 = st.radio("✉️ Select read status:", ["📩 Unread", "📖 Read"], horizontal=True) if apply_read_filter_per_app1 else "Unread"


            # 🚨 Email Importance Filter (Optional)
            st.markdown("<h4 style='color: maroon;'>🚨 Email Importance Filter</h4>", unsafe_allow_html=True)
            apply_importance_filter_per_app1 = st.checkbox("(optional) Filter by email importance?🧭")
            selected_importance_per_app1 = st.radio("🚦Select importance level:", ["🟢 Low", "🟡 Normal", "🔴 High"], horizontal=True) if apply_importance_filter_per_app1 else "Normal"

            # 🚩 Follow-Up Flag Filter (Optional)
            st.markdown("<h4 style='color: maroon;'>🚩 Follow-Up Flag Status</h4>", unsafe_allow_html=True)
            apply_flag_filter_per_app1 = st.checkbox("(optional) Filter by follow-up flag? 🏳️")
            followup_flag_status_per_app1 = st.radio("📝 Select follow-up status:", ["🚫 No Flag", "✅ Completed", "🚩 Flagged"], horizontal=True) if apply_flag_filter_per_app1 else None

            # 📏 Email Size Filter Options
            size_mode_per_app1, size_pref_per_app1, apply_size_filter_per_app1, size_cond_per_app1, size_val_per_app1, size_range_per_app1 = get_size_filter_options_per_app1()

            #==============================📎 Attachment Filters ==============================
            st.markdown("<h4 style='color: maroon;'>📎 Attachment Filters</h4>", unsafe_allow_html=True)

            # 📎 Attachment Presence Filter
            apply_attachment_filter_per_app1 = st.checkbox("(optional) Filter by attachment presence?")
            attachment_option_per_app1 = st.radio(
                "Select attachment condition:",
                ["📎 With Attachments Only", "🚫📎Without Attachments Only"],
                horizontal=True
            ) if apply_attachment_filter_per_app1 else None

            # 🧩 File Extension Filter (applies only when filtering for attachments)
            attachment_extensions_per_app1 = []
            if apply_attachment_filter_per_app1 and attachment_option_per_app1 == "📎 With Attachments Only":
                st.markdown("*(optional)* Filter specific file types (e.g., `.pdf`, `.xlsx`, `.zip`):")
                ext_input_per_app1 = st.text_input("📝 Enter comma-separated extensions:", placeholder=".pdf, .docx")
                if ext_input_per_app1:
                    attachment_extensions_per_app1 = [e.strip().lower() for e in ext_input_per_app1.split(",") if e.strip()]

            # 🏷️ Attachment Name Filter (pattern match)
            attachment_name_filter_mode_per_app1 = None
            attachment_name_filter_value_per_app1 = ""
            if apply_attachment_filter_per_app1 and attachment_option_per_app1 == "📎 With Attachments Only":
                st.markdown("*(optional)* Filter attachment name by pattern:")
                attachment_name_filter_mode_per_app1 = st.radio(
                    "🧠 Match attachment name using:",
                    ["🔍 Contains", "🚫🔍 Does not contain", "🏁 Starts with", "🏁➡️ Ends with", "🎯 Exact match"],
                    horizontal=True
                )
                attachment_name_filter_value_per_app1 = st.text_input("🔤 Enter name pattern:")

         
            # 💾 Download Settings
            download_path_per_app1, download_mode_per_app1 = get_download_settings_per_app1()

            #⬇️ Download Button & Execution
            if st.button("⬇️ Download Emails"):
                
                # 🚫 Validation Checks
                if not new_emails_per_app1:
                    st.error("❌ At least one sender email address is required.")
                elif start_date_per_app1 > end_date_per_app1:
                    st.error("❌ Start date must be before or equal to end date.")
                elif check_subject_per_app1 and not subject_value_per_app1:
                    st.error("❌ Subject filter is checked but no criteria entered.")
                else:
                    # 🚀 Begin email processing
                    start_time_per_app1 = time.time()
                    summary_per_app1, msg_count_per_app1 = process_emails_per_app1(
                        selected_folders_per_app1, new_emails_per_app1, date_range_per_app1, download_path_per_app1,
                        size_pref_per_app1, apply_size_filter_per_app1, size_cond_per_app1, size_val_per_app1, size_range_per_app1,
                        download_mode_per_app1, check_subject_per_app1, subject_filter_mode_per_app1, subject_value_per_app1, subject_case_sensitive_per_app1,
                        apply_read_filter_per_app1, read_status_per_app1,
                        apply_attachment_filter_per_app1, attachment_option_per_app1,
                        attachment_extensions_per_app1,
                        attachment_name_filter_mode_per_app1,
                        attachment_name_filter_value_per_app1,
                        apply_importance_filter_per_app1, selected_importance_per_app1,
                        apply_flag_filter_per_app1, followup_flag_status_per_app1
                    )

                    # ⏱️ Show elapsed time
                    elapsed_per_app1 = time.time() - start_time_per_app1
                    st.success(f"⏱️ Total time taken: {str(timedelta(seconds=int(elapsed_per_app1)))}")

                    # 📋 Show summary of missing emails (if any)
                    if summary_per_app1:
                        st.success(f"✅ {msg_count_per_app1} emails downloaded.")
                        st.markdown("<h4 style='color: maroon;'>📬❌ Missing Emails Summary (Preview)</h4>", unsafe_allow_html=True)
                        summary_by_date_per_app1 = defaultdict(list)
                        for item_per_app1 in summary_per_app1:
                            summary_by_date_per_app1[item_per_app1["Date"]].append(item_per_app1["Missing Email Address"])
                        for date_str_per_app1, emails_per_app1 in summary_by_date_per_app1.items():
                            st.markdown(f"**{date_str_per_app1}**")
                            for email_per_app1 in emails_per_app1:
                                st.markdown(f"- {email_per_app1}")
                    else:
                        st.success(f"✅ {msg_count_per_app1} emails downloaded. No emails missing.")

#==============================🚀 Entry Point: Run the Streamlit App ==============================
if __name__ == "__main__":
    main_per_app1()
# -------------------------------------------------------------------------------------------------
# 🎨 Bottom horizontal separator line for UI neatness
st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)
# -------------------------------------------------------------------------------------------------
