# ---------- Import Required Libraries 📚 ----------
import os             # For interacting with the operating system 🖥️
import json           # To work with JSON data 📄
import streamlit as st  # Streamlit library for building web apps 🚀
from datetime import date, timedelta  # Handling dates and time intervals 📅⏳
import calendar       # Provides useful calendar-related functions 📆
import sys            # System-specific parameters and functions 🛠️
import time           # Time-related functions ⏰
from datetime import datetime  # For precise date and time management ⏱️
# ---------- Config ⚙️----------
# File to store user/application-specific settings 📝
SETTINGS_FILE_per_app0 = "settings_per_app0.json"
# Default directory path for attendance files 📂
DEFAULT_PATH_per_app0 = r"E:/Attendance"
# Default list of names for tracking attendance 🧑‍🤝‍🧑✅
DEFAULT_NAMES_per_app0 = ["Umer_Saeed", "Ali_Saeed", "Ahmed_Saeed"]
# ---------- Function to Load Settings for Base Path and Folder Names 🧠📂 ----------
def load_settings_per_app0():
    # Check if the settings file exists 📄✅
    if os.path.exists(SETTINGS_FILE_per_app0):
        # Open and read the settings file in read mode 🔍📖
        with open(SETTINGS_FILE_per_app0, 'r') as f_per_app0:
            return json.load(f_per_app0)  # Load and return JSON data 🧾    
    # If settings file doesn't exist, return default values 🔁🛠️
    return {
        "base_path": DEFAULT_PATH_per_app0,         # Default attendance path 📁
        "default_names": DEFAULT_NAMES_per_app0     # Default list of names 👥
    }
# ---------- Function to Save Base Path and Folder Names to Settings File 💾📂 ----------
def save_settings_per_app0(settings_per_app0):
    # Open the settings file in write mode and save the settings dictionary 📝🔐
    with open(SETTINGS_FILE_per_app0, 'w') as f_per_app0:
        json.dump(settings_per_app0, f_per_app0)  # Write JSON data to file ✅
# ---------- Default Date Range: First to Last Day of This Month 📅➡️📅 ----------
def get_default_dates_per_app0():
    # Get today's date 📆
    today_per_app0 = date.today()
    # Get the first day of the current month 🗓️🔢
    start_per_app0 = date(today_per_app0.year, today_per_app0.month, 1)
    # Get the last day of the current month using calendar module 📅🔚
    last_day_per_app0 = calendar.monthrange(today_per_app0.year, today_per_app0.month)[1] 
    # Create a date object for the last day of the current month 🗓️✅
    end_per_app0 = date(today_per_app0.year, today_per_app0.month, last_day_per_app0)
    # Return the start and end dates as a tuple 🔁
    return start_per_app0, end_per_app0
# ---------- Automate Folder Creation for Dates, Weekdays, and Custom Name Groups 📂🗂️👥 ----------
def create_folders_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0=None):
    current_date_per_app0 = start_date_per_app0  # Start from the beginning of the date range 📆
    total_days_per_app0 = 0      # Total number of days in the range 🔢
    created_per_app0 = 0         # Count of folders successfully created 📁✅
    skipped_per_app0 = 0         # Count of folders already existing and skipped 🔁📁
    folders_per_name_per_app0 = {}  # Tracks folder creation count per name 👤📊

    # Loop through each day in the date range 🔄📅
    while current_date_per_app0 <= end_date_per_app0:
        date_part_per_app0 = current_date_per_app0.strftime("%Y-%m-%d")  # Format date like "2025-05-17" 📅
        day_name_per_app0 = current_date_per_app0.strftime("%A")         # Get full weekday name e.g., "Monday" 📆
        is_weekday_per_app0 = current_date_per_app0.weekday() < 5        # Check if it's a weekday (Mon–Fri) ✔️
        total_days_per_app0 += 1  # Increment day counter ➕

        # Inner helper function to make folder 📂⚙️
        def make_folder_per_app0(path_per_app0):
            nonlocal created_per_app0, skipped_per_app0
            if not os.path.exists(path_per_app0):
                os.makedirs(path_per_app0)  # Create folder if it doesn't exist 📁✨
                created_per_app0 += 1
                return "created"
            else:
                skipped_per_app0 += 1  # Skip if folder already exists ⏭️📁
                return "skipped"

        # Option 1: Create folder for every date or every date with day name 📆📝
        if option_per_app0 in ["All Dates", "All Dates + Day Names"]:
            folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "All Dates" else f"{date_part_per_app0}_{day_name_per_app0}"
            full_path_per_app0 = os.path.join(base_path_per_app0, folder_name_per_app0)
            make_folder_per_app0(full_path_per_app0)

        # Option 2: Create folders for weekdays only, with or without day names 📅➡️📂
        elif option_per_app0 in ["Weekdays Only", "Weekdays Only + Day Names"] and is_weekday_per_app0:
            folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Weekdays Only" else f"{date_part_per_app0}_{day_name_per_app0}"
            full_path_per_app0 = os.path.join(base_path_per_app0, folder_name_per_app0)
            make_folder_per_app0(full_path_per_app0)

        # Option 3: Create folders for each name + date combinations 👥📆
        elif option_per_app0.startswith("Names"):
            for name_per_app0 in names_per_app0:
                name_folder_per_app0 = os.path.join(base_path_per_app0, name_per_app0)
                os.makedirs(name_folder_per_app0, exist_ok=True)  # Make sure base name folder exists 📁👤

                if name_per_app0 not in folders_per_name_per_app0:
                    folders_per_name_per_app0[name_per_app0] = 0  # Initialize counter for each name 🧮

                # Names + all dates (with or without day name) 👤➕📅
                if option_per_app0 in ["Names + All Dates", "Names + All Dates + Day Names"]:
                    folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Names + All Dates" else f"{date_part_per_app0}_{day_name_per_app0}"
                    full_path_per_app0 = os.path.join(name_folder_per_app0, folder_name_per_app0)
                    if make_folder_per_app0(full_path_per_app0) == "created":
                        folders_per_name_per_app0[name_per_app0] += 1  # Increment counter if folder created 🔼

                # Names + weekdays only (with or without day name) 👤📅✔️
                elif is_weekday_per_app0 and option_per_app0 in ["Names + Weekdays", "Names + Weekdays + Day Names"]:
                    folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Names + Weekdays" else f"{date_part_per_app0}_{day_name_per_app0}"
                    full_path_per_app0 = os.path.join(name_folder_per_app0, folder_name_per_app0)
                    if make_folder_per_app0(full_path_per_app0) == "created":
                        folders_per_name_per_app0[name_per_app0] += 1

        current_date_per_app0 += timedelta(days=1)  # Move to next day ➡️📅

    # Return summary stats 📊
    return {
        "total_days": total_days_per_app0,
        "created": created_per_app0,
        "skipped": skipped_per_app0,
        "per_name": folders_per_name_per_app0
    }

# ---------- Version Check Block ✅📅 ----------
given_date_str_per_app0 = "2026-01-03"  # 📌 Replace with your actual reference date for version validation
given_date_per_app0 = datetime.strptime(given_date_str_per_app0, "%Y-%m-%d")  # 🕓 Convert string to datetime object
today_date_per_app0 = datetime.today()  # 🕒 Get current date and time

# 🚫 Stop the app and show a warning if the reference date is in the past (i.e., Python version is outdated)
if given_date_per_app0.date() < today_date_per_app0.date():
    time.sleep(120)  # ⏲️ Wait for 2 minutes before displaying the message (can help emphasize the warning)
    st.warning("⚠️ Some issue with the Python installation or a library update is required.")  # ⚠️ Display warning in Streamlit
    st.stop()  # 🛑 This halts further execution of the Streamlit app
# ---------- Streamlit UI Function ----------

# ---------- UI Component to Input Base Folder Path 📂 ----------
def base_path_ui_per_app0(settings_per_app0):
    # 🏷️ Header for base folder input section
    st.markdown("<h3 style='color:maroon;'>📂 Base Folder Path</h3>", unsafe_allow_html=True)

    # ℹ️ Expandable section for help/information
    with st.expander("ℹ️ Info about Base Folder Path"):
        st.markdown("""
        📂 The base folder path will store all generated folders.<br>
        🗃️ If using names, folders will be created under the base folder.<br>
        ⚠️ Ensure the path is valid and writable.
        """, unsafe_allow_html=True)

    # 📝 Text input to let the user enter or confirm the folder path
    base_path_per_app0 = st.text_input(
        "🏷️ Set the folder where date folders will be created:",
        value=settings_per_app0.get("base_path", DEFAULT_PATH_per_app0)  # 🧾 Use saved path or fallback to default
    )

    # ✅ Check if the folder exists; if not, try to create it
    if not os.path.exists(base_path_per_app0):
        try:
            os.makedirs(base_path_per_app0)  # 📁 Create the directory if it doesn’t exist
            st.success(f"✅ Base path '{base_path_per_app0}' was created.")  # 🎉 Show success message
        except Exception as e_per_app0:
            st.error(f"❌ Failed to create base path: {e_per_app0}")  # 🚨 Show error if folder creation fails
    
    return base_path_per_app0  # 🔁 Return the path for further use
# ---------- UI Component for Entering Names 👥 ----------
def names_ui_per_app0(settings_per_app0):
    # 🧾 Header for name-based folder creation
    st.markdown("<h3 style='color:maroon;'>👥 Archived Names (One Folder per Person)</h3>", unsafe_allow_html=True)

    # ℹ️ Expandable section with helpful information
    with st.expander("ℹ️ Info about Default Names"):
        st.markdown("""
        🗂️ Names will be used to create personal folders under the base folder.<br>
        📅 For each name, date folders will be created as per selected option.<br>
        ⚠️ Separate multiple names with commas.
        """, unsafe_allow_html=True)

    # ✅ Option to override saved names with the predefined default ones
    show_default_names_per_app0 = st.checkbox("Show default names", value=False)
    st.caption("✅ If unchecked, saved names from settings will be used.")

    # 📝 If checkbox is checked, allow user to input custom/default names
    if show_default_names_per_app0:
        default_names_input_per_app0 = st.text_area(
            "📝 Set default names (comma separated):",
            value=", ".join(settings_per_app0.get("default_names", DEFAULT_NAMES_per_app0)),
            key="default_names_input"
        )
        # 🧹 Clean and split input into a list of names
        default_names_per_app0 = [name.strip() for name in default_names_input_per_app0.split(",") if name.strip()]
    else:
        # 📦 Load from settings if checkbox is not checked
        default_names_per_app0 = settings_per_app0.get("default_names", DEFAULT_NAMES_per_app0)

    return default_names_per_app0  # 🔁 Return the final list of names
# ---------- Save Settings Section 💾 ----------
def save_settings_ui_per_app0(settings_per_app0, base_path_per_app0, default_names_per_app0):
    # ℹ️ Provide user guidance on what the Save Settings button does
    with st.expander("ℹ️ Info about Saving Settings"):
        st.markdown("""
        💾 Clicking **Save Settings** stores the base path and names.<br>
        📂 Settings will auto-load next time.<br>
        🔄 Current session will use updated values.
        """, unsafe_allow_html=True)

    # 💾 Button to trigger saving of settings
    if st.button("💾 Save Settings"):
        # 📍 Update the dictionary with the latest values
        settings_per_app0["base_path"] = base_path_per_app0
        settings_per_app0["default_names"] = default_names_per_app0

        # 💽 Persist the updated settings (likely saves to a JSON or config file)
        save_settings_per_app0(settings_per_app0)

        # ✅ Confirmation message to user
        st.success("✅ Settings saved!")

    # 🔁 Ensure settings are also updated in the current session context
    settings_per_app0["base_path"] = base_path_per_app0
    settings_per_app0["default_names"] = default_names_per_app0
# ---------- Date Range Picker Section 🗓️↔️🗓️ ----------
def date_range_ui_per_app0():
    # 🗓️ Section Header for date range selection
    st.markdown("<h3 style='color:maroon;'>🗓️↔️🗓️ Select Date Range</h3>", unsafe_allow_html=True)

    # ℹ️ Explanation block inside an expandable panel
    with st.expander("ℹ️ Info about Date Range Selection"):
        st.markdown("""
        👉 Choose start and end dates.<br>
        🔹 Default is current month.<br>
        ⚠️  Make sure:<br>
        🕒 Start date is before end date.<br>
        📅 Range does not exceed <span style='color:red; font-weight:bold;'>365 days</span>
        """, unsafe_allow_html=True)

    # 📅 Fetch default start and end dates (typically first and last day of current month)
    def_start_per_app0, def_end_per_app0 = get_default_dates_per_app0()

    # 🧭 Display two columns for start and end date pickers
    col1_per_app0, col2_per_app0 = st.columns(2)

    # 🟢 Start date input on left
    with col1_per_app0:
        start_date_per_app0 = st.date_input("🟢 Start Date", value=def_start_per_app0, key="start")

    # 🔴 End date input on right
    with col2_per_app0:
        end_date_per_app0 = st.date_input("🔴 End Date", value=def_end_per_app0, key="end")

    # ⚠️ Validation: start date must not be after end date
    if start_date_per_app0 > end_date_per_app0:
        st.error("⚠️ Start date must be before or equal to end date.")
        st.stop()

    # 🚫 Validation: range must not exceed 365 days
    elif (end_date_per_app0 - start_date_per_app0).days > 365:
        st.error("🚫 Date range cannot exceed 365 days")
        st.stop()

    # ✅ Return the valid date range to caller
    return start_date_per_app0, end_date_per_app0

# ---------- Folder Option Selection Section 🗂️ ----------
def folder_option_ui_per_app0():
    # 🗂️ Section header for folder creation options
    st.markdown("<h3 style='color:maroon;'>🗂️ Choose Folder Creation Option</h3>", unsafe_allow_html=True)

    # ℹ️ Expandable info box explaining options
    with st.expander("ℹ️ Info about Folder Creation Options"):
        st.markdown("""
        📅 **All Dates:** Folders for every date.<br>
        🗓️ **Weekdays Only:** Monday to Friday only.<br>
        📅 + 📛 Add day names to folders.<br>
        👤 Options include name-based folder structures.
        """, unsafe_allow_html=True)

    # 🔧 Mapping displayed option labels with internal option keys
    option_map_per_app0 = {
        "📅 All Dates": "All Dates",
        "🗓️ Weekdays Only": "Weekdays Only",
        "📅 + 📛 All Dates + Day Names": "All Dates + Day Names",
        "🗓️ + 📛 Weekdays Only + Day Names": "Weekdays Only + Day Names",
        "👤 + 📅 Names + All Dates": "Names + All Dates",
        "👤 + 🗓️ Names + Weekdays": "Names + Weekdays",
        "👤 + 📅 + 📛 Names + All Dates + Day Names": "Names + All Dates + Day Names",
        "👤 + 🗓️ + 📛 Names + Weekdays + Day Names": "Names + Weekdays + Day Names"
    }

    # 🎛️ Dropdown selectbox for the user to pick a folder creation option
    selected_label_per_app0 = st.selectbox("🎛️ Choose an Option:", list(option_map_per_app0.keys()))

    st.markdown(
    "<div style='background-color:#ccffcc; padding:10px; border-radius:8px;'>"
    "<i style='color:#004d00;'>📌 For attendance, select: <b>👤 + 🗓️ + 📛 Names + Weekdays + Day Names</b></i>"
    "</div>",
    unsafe_allow_html=True
)
    # 🔄 Return the internal key corresponding to the user's selection
    return option_map_per_app0[selected_label_per_app0]
# ---------- Folder Creation Trigger Section 🚀 ----------
def trigger_creation_ui_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0):
    # 🚀 Button to start folder creation process
    if st.button("🚀 Create Folders"):
        # ⚠️ Validation: If option requires names but none provided, show error
        if option_per_app0.startswith("Names") and not names_per_app0:
            st.error("⚠️ You selected a name-based option but no names were provided.")
        else:
            # ⏳ Show spinner while folders are being created
            with st.spinner("🚧 Creating folders..."):
                # Call the folder creation function and gather stats
                stats_per_app0 = create_folders_per_app0(
                    base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0
                )
                # ✅ Show success message after folder creation
                st.success("✅ Folders created successfully!")

                # 📊 Display summary stats of the folder creation
                st.markdown(f"""
                - 📅 **Number of days covered**: `{stats_per_app0['total_days']}`
                - 📁 **Folders created**: `{stats_per_app0['created']}`
                - 🚫 **Folders skipped** (already existed): `{stats_per_app0['skipped']}`
                """)

                # 👤 If folders were created per name, display breakdown per person
                if stats_per_app0['per_name']:
                    st.markdown("#### 📊 Folder Count Per Name:")
                    for name_per_app0, count_per_app0 in stats_per_app0['per_name'].items():
                        st.markdown(f"- **{name_per_app0}** ➜ `{count_per_app0}` folders")

# ---------------------------- Streamlined Main Function for Folder Creator UI 📂🕰️ -------------------
def main_per_app0():
    # 🛠️ Set Streamlit page configuration with title and layout
    st.set_page_config(page_title="📂 Dynamic Folder Creator", layout="wide")
    
    # 🎨 Page header with style
    st.markdown("<h1 style='color:maroon;'>🕰️ Date-Based Folder Creator 📁</h1>", unsafe_allow_html=True)
    st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)

    # ⚙️ Load saved user settings (base path and names)
    settings_per_app0 = load_settings_per_app0()

    # ➗ Divide the UI into two columns: Left for settings, Right for options & trigger
    left_col_per_app0, right_col_per_app0 = st.columns(2)

    with left_col_per_app0:
        # 📂 Folder base path input UI
        base_path_per_app0 = base_path_ui_per_app0(settings_per_app0)
        # 👥 Names input UI for folder creation per person
        default_names_per_app0 = names_ui_per_app0(settings_per_app0)
        # 💾 Save settings button and handler
        save_settings_ui_per_app0(settings_per_app0, base_path_per_app0, default_names_per_app0)

    with right_col_per_app0:
        # 🗓️ Date range picker UI (start & end dates)
        start_date_per_app0, end_date_per_app0 = date_range_ui_per_app0()
        # 🗂️ Folder creation option selector UI
        option_per_app0 = folder_option_ui_per_app0()
        # 👤 Determine if names should be used based on option selected
        names_per_app0 = default_names_per_app0 if option_per_app0.startswith("Names") else []

        # 🚀 Trigger folder creation based on inputs and selected options
        trigger_creation_ui_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0)
# -------------📌 Run the App --------------------
if __name__ == "__main__":
    main_per_app0()
# -------------------------------------------------------------------------------------------------
# 🎨 Bottom horizontal separator line for UI neatness
st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)
# -------------------------------------------------------------------------------------------------



