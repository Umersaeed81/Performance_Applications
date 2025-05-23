# ------------📦🔁 Import Required Libraries --------------------------------------#
import os  # 📁 OS-level file and directory operations
import json  # 📄🔧 Read and write JSON configuration files
import pythoncom  # 🧠 Required to initialize COM libraries in multi-threaded environments
import win32com.client  # 🤖 Automate and interact with Outlook and other Microsoft apps via COM
import streamlit as st  # 🌐 Build interactive web apps with Python
import pandas as pd  # 🐼 Handle tabular data and perform data analysis
from datetime import datetime  # ⏰ Work with date and time objects
import glob  # 🔍 Perform pattern-based file searching
import time  # ⏳ Track execution time for performance monitoring
from collections import defaultdict  # 📦 Efficient grouping of data by keys
import collections  # 🧺 General-purpose data structures (namedtuple, deque, etc.)
import datetime  # 📅 Alternative access to date and time functions
from datetime import datetime, timedelta
#------------ 📬📦 Path to Store Extracted Emails in JSON Format------------
EMAIL_JSON_app1 = "settings_per_app1.json"  # 💾 Emails will be saved here for later use or review
#------------  ⚙️🖥️ Streamlit Page Configuration for Outlook Email Downloader------------ 
def configure_page_app1():
    # 🧾 Set the page title and layout style (wide layout for better display)
    st.set_page_config(page_title="Outlook Email Downloader", layout="wide")
    # 🏷️ Add a custom HTML header with maroon color and email icon
    st.markdown("<h1 style='color: maroon;'>📥 Outlook Email Downloader (Smallest Email per Day per Sender)</h1>", unsafe_allow_html=True)
    # 📏 Add a stylish horizontal line as a divider
    st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)
#------------ 📧🔌 Initialize Outlook Application for Email Access------------
def initialize_outlook_app1():
    pythoncom.CoInitialize()  # ⚙️ Initialize COM library (important for multithreaded environments)
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # 🗂️ Connect to Outlook's MAPI namespace to access emails
#------------ 🗂️📂 PST Selection Logic – Choose Between Current Outlook or External PST File------------
def select_pst_option_app1(outlook_app1):
    # 🔘 User selects how to access emails: from current Outlook profile or an external PST file
    # st.markdown("<h4 style='color: maroon;'>📌 Select PST Option</h4>", unsafe_allow_html=True)

    use_current_pst_app1 = st.checkbox("💼 Use current Outlook PST", value=False)  # ✅ Option to use currently loaded Outlook profile
    browse_pst_file_app1 = st.checkbox("🧭 Browse PST file", value=False)  # 📁 Option to load a specific PST file from disk

    root_folder_app1 = None  # 📌 Will hold the root folder of selected PST

    # 🚫 Prevent user from selecting both options at once
    if use_current_pst_app1 and browse_pst_file_app1:
        st.warning("⚠️ Please select only one PST option.")
    elif not use_current_pst_app1 and not browse_pst_file_app1:
        st.info("ℹ️ Please select a PST option to proceed.")
    else:
        if browse_pst_file_app1:
            # 📍 Prompt for PST file path
            pst_path_app1 = st.text_input("Enter full path to PST file (e.g., E:\\Email\\LTE_KPI_REPORTING.pst)")
            if pst_path_app1:
                if not os.path.exists(pst_path_app1):
                    st.error("❌ PST file not found.")
                else:
                    try:
                        # ➕ Add the external PST file to Outlook
                        outlook_app1.AddStore(pst_path_app1)

                        # 🔍 Try to find the newly added PST's root folder
                        for i_app1 in range(outlook_app1.Folders.Count):
                            folder_app1 = outlook_app1.Folders.Item(i_app1 + 1)
                            if folder_app1.FolderPath.endswith(os.path.basename(pst_path_app1).replace(".pst", "")):
                                root_folder_app1 = folder_app1
                                break

                        # ❌ If PST is added but root folder not detected
                        if not root_folder_app1:
                            st.error("❌ Could not locate the root folder for the provided PST path.")
                    except Exception as e_app1:
                        st.error(f"❌ Error loading PST: {e_app1}")
        elif use_current_pst_app1:
            # 📥 Use default Inbox folder from current Outlook profile
            root_folder_app1 = outlook_app1.GetDefaultFolder(6)  # 6 refers to the Inbox
    
    return root_folder_app1  # 📤 Return the selected root folder for further email processing
#----------- 📁🔍 Recursively Retrieve All Subfolders from the Selected Outlook Folder-----------
def get_all_folders_app1(folder_app1, path_app1=""):
    # 🧩 Build the full path by appending the current folder name
    full_path_app1 = os.path.join(path_app1, folder_app1.Name) if path_app1 else folder_app1.Name
    
    # 📦 Initialize list with the current folder's full path and object
    folders_app1 = [(full_path_app1, folder_app1)]
    
    # 🔄 Loop through all subfolders and collect them recursively
    for sub_app1 in folder_app1.Folders:
        folders_app1.extend(get_all_folders_app1(sub_app1, full_path_app1))  # 📂 Dive deeper into subfolders
    
    return folders_app1  # 📤 Return a list of tuples: (folder_path, folder_object)
#----------- 📂✅ Folder Selection UI – Choose Folders to Scan for Emails & Attachments-----------
def get_folder_selections_app1(root_folder_app1):
    all_folders_app1 = get_all_folders_app1(root_folder_app1)  # 📋 Get all folders and subfolders from the selected PST or Inbox

    # 📝 Optional section title (commented out)
    # st.markdown("<h4 style='color: maroon;'>🗃️ Filter Folders for Attachment Extraction</h4>", unsafe_allow_html=True)
    
    # 💡 Helpful tip for users
    st.markdown("> ✅ **Tip:** Select one or more folders to scan for emails.")

    selected_app1 = []  # 📌 List to store selected folders
    
    # 🔘 Show checkboxes for each folder and collect user selections
    for folder_path_app1, folder_obj_app1 in all_folders_app1:
        if st.checkbox(folder_path_app1, value=False):  # 📎 User selects this folder
            selected_app1.append((folder_path_app1, folder_obj_app1))
    
    return selected_app1  # 📤 Return list of selected folders for further processing
#----------- 📤📧 Retrieve the Sender's SMTP Email Address from an Outlook Message-----------
def get_sender_smtp_address_app1(message_app1):
    try:
        # ✉️ If sender email type is SMTP, return it directly
        if message_app1.SenderEmailType == "SMTP":
            return message_app1.SenderEmailAddress or ""

        exch_user_app1 = None

        # 🏢 Try to get Exchange user details if email type is not SMTP
        try:
            exch_user_app1 = message_app1.Sender.GetExchangeUser()
        except Exception:
            pass  # ⚠️ Gracefully handle cases where exchange user info is unavailable

        # 📬 Return the primary SMTP address from Exchange user, if available
        if exch_user_app1:
            smtp_app1 = exch_user_app1.PrimarySmtpAddress
            if smtp_app1:
                return smtp_app1

        # 🔁 Fallback to sender's email address if SMTP not found via Exchange
        return message_app1.SenderEmailAddress or ""
    
    except Exception:
        # ❌ Return empty string if any unexpected error occurs
        return ""
#----------- 💾📨 Save Outlook Email as .msg File (with Safe Filename)-----------
def save_mail_as_msg_app1(mail_app1, folder_path_app1, filename_app1):
    # 🧹 Clean the filename by removing characters not allowed in Windows filenames
    safe_filename_app1 = "".join(c for c in filename_app1 if c not in r'\\/:*?"<>|')

    # 📁 Construct full file path for saving
    filepath_app1 = os.path.join(folder_path_app1, safe_filename_app1)

    try:
        # 💌 Save the email as a .msg file (3 = Outlook format)
        mail_app1.SaveAs(filepath_app1, 3)
        return True  # ✅ Save successful
    except Exception:
        return False  # ❌ Save failed (e.g., permission issue, invalid path)
# ----------- 📥 Fetch, Save, and Report Outlook Emails from Selected Folders -----------
def process_emails_app1(selected_folders_app1, sender_emails_app1, date_range_app1, download_path_app1, download_mode_app1="Delete all files if exist."):
    with st.spinner("⏳ Processing emails..."):
        # 📁 Create download directory if it doesn't exist
        if not os.path.exists(download_path_app1):
            os.makedirs(download_path_app1)

        # 🧹 Clean directory if mode is set to delete existing files
        if download_mode_app1 == "Delete all files if exist.":
            files_app1 = glob.glob(os.path.join(download_path_app1, '*'))
            for f_app1 in files_app1:
                try:
                    if os.path.isfile(f_app1):
                        os.remove(f_app1)
                except Exception as e_app1:
                    st.error(f"❌ Failed to delete {f_app1}: {e_app1}")

        # 🗂️ Dictionary to group emails by (date, sender)
        emails_by_date_sender = collections.defaultdict(list)

        # 📆 Convert date range into datetime objects
        start_date = min(date_range_app1)
        end_date = max(date_range_app1)
        start_dt = datetime.combine(start_date, datetime.min.time())
        end_dt = datetime.combine(end_date, datetime.max.time())

        # 📬 Loop through each selected folder and fetch emails
        for _, folder_app1 in selected_folders_app1:
            items_app1 = folder_app1.Items
            restriction_app1 = f"[ReceivedTime] >= '{start_dt.strftime('%m/%d/%Y %H:%M %p')}' AND [ReceivedTime] <= '{end_dt.strftime('%m/%d/%Y %H:%M %p')}'"
            try:
                restricted_items_app1 = items_app1.Restrict(restriction_app1)
            except Exception:
                restricted_items_app1 = items_app1

            # 📩 Loop through each email item in the filtered list
            for i_app1 in range(restricted_items_app1.Count):
                try:
                    mail_app1 = restricted_items_app1.Item(i_app1 + 1)
                    if mail_app1.Class != 43:  # 43 = MailItem
                        continue
                    smtp_app1 = get_sender_smtp_address_app1(mail_app1)
                    if smtp_app1.lower() in [email.lower() for email in sender_emails_app1]:
                        received_date = mail_app1.ReceivedTime.date()
                        emails_by_date_sender[(received_date, smtp_app1.lower())].append(mail_app1)
                except Exception:
                    continue

        summary_app1 = []         # 📝 Summary of missing emails
        msg_counter_app1 = 0      # 🔢 Counter for saved emails

        # 📅 Loop through each date in range
        for day_app1 in date_range_app1:
            missing_emails_app1 = []

            # 📧 Check emails for each sender on the specific date
            for email_app1 in sender_emails_app1:
                key = (day_app1.date(), email_app1.lower())
                mails = emails_by_date_sender.get(key, [])

                if mails:
                    # ✅ Pick the smallest size email for saving
                    smallest_mail_app1 = min(mails, key=lambda m: m.Size)
                    sender_name_app1 = smallest_mail_app1.SenderName or "UnknownSender"
                    subject_app1 = (smallest_mail_app1.Subject or "NoSubject")[:50]
                    filename_app1 = f"{day_app1.strftime('%Y-%m-%d')} - {sender_name_app1} - {subject_app1}.msg"
                    safe_filename_app1 = "".join(c for c in filename_app1 if c not in r'\\/:*?"<>|')
                    filepath_app1 = os.path.join(download_path_app1, safe_filename_app1)

                    # 📝 Rename file if overwrite is disabled
                    if download_mode_app1 == "Do not overwrite, rename it.":
                        filepath_app1 = get_unique_filename_app1(filepath_app1)

                    try:
                        # 💾 Save the email to the disk
                        smallest_mail_app1.SaveAs(filepath_app1, 3)
                        msg_counter_app1 += 1
                    except Exception as e_app1:
                        st.error(f"❌ Failed to save email: {e_app1}")
                else:
                    # ⚠️ Email not found for this sender on this day
                    missing_emails_app1.append(email_app1)

            # ➕ Add missing emails to the summary list
            for missing_app1 in missing_emails_app1:
                summary_app1.append({
                    "Date": day_app1.strftime('%Y-%m-%d'),
                    "Missing Email Address": missing_app1
                })

        # 📊 Export summary report to Excel
        xlsx_path_app1 = os.path.join(download_path_app1, "Missing_Email_Summary.xlsx")
        df_app1 = pd.DataFrame(summary_app1)
        df_app1.to_excel(xlsx_path_app1, index=False)

    # 🔚 Return missing summary and count of saved messages
    return summary_app1, msg_counter_app1

#----------- 💾📂 Load Previously Saved Emails from JSON File-----------
#EMAIL_JSON_app1 = "saved_emails_app1.json"  # Assuming global variable for saved emails filename
EMAIL_JSON_app1 = "settings_per_app1.json"  # 💾 Emails will be saved here for later use or review
def load_saved_emails_app1():
    # 🔍 Check if the saved emails JSON file exists
    if os.path.exists(EMAIL_JSON_app1):
        # 📖 Read and load JSON data
        with open(EMAIL_JSON_app1, 'r') as f_app1:
            return json.load(f_app1)
    # 🚫 Return empty list if no saved emails found
    return []

#----------- 💾📝 Save Emails Data to JSON File for Persistent Storage-----------
def save_emails_to_json_app1(emails_app1):
    # 📂 Open JSON file in write mode and save email data
    with open(EMAIL_JSON_app1, 'w') as f_app1:
        json.dump(emails_app1, f_app1)

#----------- ⚙️📂 Configure Download Settings and Folder Selection UI-----------
def get_download_settings_app1(default_download_path_app1="D:\\Downloaded_Attachments"):
    # 📂 Prompt user to select or enter download folder path
    st.markdown("<h4 style='color: maroon;'>📂 Select Download Location</h4>", unsafe_allow_html=True)
    download_path_app1 = st.text_input("📂 Download Folder Path:", value=default_download_path_app1)
    os.makedirs(download_path_app1, exist_ok=True)

    # 🛠️ Allow user to configure how files are handled during download
    st.markdown("<h4 style='color: maroon;'>🛠️ Download Mode Configuration</h4>", unsafe_allow_html=True)
    download_mode_app1 = st.selectbox("Choose download mode ⚙️:", [
        "Delete all files if exist.",
        "Do not overwrite, rename it."
    ])

    # 🔄 Return selected folder path and download mode
    return download_path_app1, download_mode_app1


#-----------  📝🔄 Generate Unique Filename to Avoid Overwriting Existing Files----------- 
def get_unique_filename_app1(file_path_app1):
    base_app1, ext_app1 = os.path.splitext(file_path_app1)
    counter_app1 = 1
    new_path_app1 = file_path_app1
    # 🔁 Loop until a unique filename is found
    while os.path.exists(new_path_app1):
        new_path_app1 = f"{base_app1}_{counter_app1}{ext_app1}"
        counter_app1 += 1
    # ✅ Return the unique filename
    return new_path_app1

#----------- 🚀🖥️ Main Application Logic for Outlook Email Downloader-----------
def main_app1():


    # ⚙️ Setup Streamlit page configuration and UI
    configure_page_app1()
    
    # 📧 Initialize Outlook COM namespace
    outlook_app1 = initialize_outlook_app1()

    # 🗂️ Create two-column layout for inputs and controls
    col1_app1, col2_app1 = st.columns(2)

    with col1_app1:
        # 📌 PST selection options UI
        st.markdown("<h4 style='color: maroon;'>📌 Select PST Option</h4>", unsafe_allow_html=True)
        root_folder_app1 = select_pst_option_app1(outlook_app1)

        if root_folder_app1:
            # 🗃️ Folder selection UI to choose folders to scan
            st.markdown("<h4 style='color: maroon;'>🗃️ Select Folder(s) to Scan for Emails</h4>", unsafe_allow_html=True)
            selected_folders_app1 = get_folder_selections_app1(root_folder_app1)
        else:
            selected_folders_app1 = []

    with col2_app1:
        # 💾 Load saved sender email addresses and provide input UI
        saved_emails_app1 = load_saved_emails_app1()
        st.markdown("<h4 style='color: maroon;'>🆔 Sender Email Address</h4>", unsafe_allow_html=True)
        email_input_app1 = st.text_area("✉️ Enter Sender Email Addresses (comma separated):", ", ".join(saved_emails_app1))
        new_emails_app1 = [e_app1.strip() for e_app1 in email_input_app1.split(',') if e_app1.strip()]

        # 💾 Button to save entered sender email addresses
        if st.button("💾 Save Email Addresses"):
            save_emails_to_json_app1(new_emails_app1)
            st.success("✅ Email addresses saved!")

        # 📆 Date range selection UI
        st.markdown("<h4 style='color: maroon;'>📆 Select Date Range</h4>", unsafe_allow_html=True)
        d1_app1, d2_app1 = st.columns(2)
        with d1_app1:
            start_date_app1 = st.date_input("🟢📆 Start Date", key="start_date_app1")
        with d2_app1:
            end_date_app1 = st.date_input("🔴📆 End Date", key="end_date_app1")
        date_range_app1 = pd.date_range(start=start_date_app1, end=end_date_app1).to_pydatetime().tolist()

        # 📂 Download path and mode configuration UI
        download_path_app1, download_mode_app1 = get_download_settings_app1()

        # ⬇️ Trigger email processing and downloading
        if st.button("⬇️ Download Smallest Emails"):
            # ⚠️ Validate inputs before processing
            if not selected_folders_app1:
                st.warning("⚠️ Please select at least one folder to scan.")
            elif not new_emails_app1:
                st.error("❌ At least one sender email address is required.")
            elif start_date_app1 > end_date_app1:
                st.error("❌ Start date must be before or equal to end date.")
            else:
                start_time = time.time()

                # 📧 Process emails with selected settings
                summary_app1, msg_counter_app1 = process_emails_app1(
                    selected_folders_app1,
                    new_emails_app1,
                    date_range_app1,
                    download_path_app1,
                    download_mode_app1  
                )

                elapsed_time = time.time() - start_time
                #formatted_time = str(datetime.timedelta(seconds=int(elapsed_time)))
                formatted_time = str(timedelta(seconds=int(elapsed_time)))
                st.success(f"⏱️ Total time taken: {formatted_time}")

                if summary_app1:
                    st.success(f"✅ {msg_counter_app1} emails downloaded. Summary saved to: {download_path_app1}\\missing_email_summary.xlsx")
                    st.markdown("<h4 style='color: maroon;'>📬❌ Missing Emails Summary</h4>", unsafe_allow_html=True)

                    # Group missing emails by date
                    summary_by_date = defaultdict(list)
                    for item in summary_app1:
                        summary_by_date[item["Date"]].append(item["Missing Email Address"])

                    for date_str, emails in summary_by_date.items():
                        st.markdown(f"**{date_str}**")
                        for email in emails:
                            st.markdown(f"- {email}")
                else:
                    st.success(f"✅ No missing emails. All emails found and downloaded. Total downloaded: {msg_counter_app1}")


# -----------🏁🚀 Entry Point of the Application-----------
if __name__ == "__main__":
    # ▶️ Run the main function to start the Streamlit app
    main_app1()
#-------------------------------------------------------------
st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)
#-------------------------------------------------------------
