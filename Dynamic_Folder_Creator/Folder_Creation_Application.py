# ---------- Import Required Libraries ----------
import os
import json
import streamlit as st
from datetime import date, timedelta
import calendar
import sys
import time
from datetime import datetime
# ---------- Config ----------
SETTINGS_FILE_per_app0 = "settings.json"
DEFAULT_PATH_per_app0 = r"E:/Attendance"
DEFAULT_NAMES_per_app0 = ["Umer", "Ali", "Ahmed"]

# ---------- Function to Load Settings for Base Path and Folder Names ----------
def load_settings_per_app0():
    if os.path.exists(SETTINGS_FILE_per_app0):
        with open(SETTINGS_FILE_per_app0, 'r') as f_per_app0:
            return json.load(f_per_app0)
    return {"base_path": DEFAULT_PATH_per_app0, "default_names": DEFAULT_NAMES_per_app0}

# ---------- Function to Save Base Path and Folder Names to Settings File ----------
def save_settings_per_app0(settings_per_app0):
    with open(SETTINGS_FILE_per_app0, 'w') as f_per_app0:
        json.dump(settings_per_app0, f_per_app0)

# ---------- Default Date Range: First to Last Day of This Month ----------
def get_default_dates_per_app0():
    today_per_app0 = date.today()
    start_per_app0 = date(today_per_app0.year, today_per_app0.month, 1)
    last_day_per_app0 = calendar.monthrange(today_per_app0.year, today_per_app0.month)[1]
    end_per_app0 = date(today_per_app0.year, today_per_app0.month, last_day_per_app0)
    return start_per_app0, end_per_app0

# ---------- Automate Folder Creation for Dates, Weekdays, and Custom Name Groups ----------
def create_folders_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0=None):
    current_date_per_app0 = start_date_per_app0
    total_days_per_app0 = 0
    created_per_app0 = 0
    skipped_per_app0 = 0
    folders_per_name_per_app0 = {}

    while current_date_per_app0 <= end_date_per_app0:
        date_part_per_app0 = current_date_per_app0.strftime("%Y-%m-%d")
        day_name_per_app0 = current_date_per_app0.strftime("%A")
        is_weekday_per_app0 = current_date_per_app0.weekday() < 5
        total_days_per_app0 += 1

        def make_folder_per_app0(path_per_app0):
            nonlocal created_per_app0, skipped_per_app0
            if not os.path.exists(path_per_app0):
                os.makedirs(path_per_app0)
                created_per_app0 += 1
                return "created"
            else:
                skipped_per_app0 += 1
                return "skipped"

        if option_per_app0 in ["All Dates", "All Dates + Day Names"]:
            folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "All Dates" else f"{date_part_per_app0}_{day_name_per_app0}"
            full_path_per_app0 = os.path.join(base_path_per_app0, folder_name_per_app0)
            make_folder_per_app0(full_path_per_app0)

        elif option_per_app0 in ["Weekdays Only", "Weekdays Only + Day Names"] and is_weekday_per_app0:
            folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Weekdays Only" else f"{date_part_per_app0}_{day_name_per_app0}"
            full_path_per_app0 = os.path.join(base_path_per_app0, folder_name_per_app0)
            make_folder_per_app0(full_path_per_app0)

        elif option_per_app0.startswith("Names"):
            for name_per_app0 in names_per_app0:
                name_folder_per_app0 = os.path.join(base_path_per_app0, name_per_app0)
                os.makedirs(name_folder_per_app0, exist_ok=True)

                if name_per_app0 not in folders_per_name_per_app0:
                    folders_per_name_per_app0[name_per_app0] = 0

                if option_per_app0 in ["Names + All Dates", "Names + All Dates + Day Names"]:
                    folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Names + All Dates" else f"{date_part_per_app0}_{day_name_per_app0}"
                    full_path_per_app0 = os.path.join(name_folder_per_app0, folder_name_per_app0)
                    if make_folder_per_app0(full_path_per_app0) == "created":
                        folders_per_name_per_app0[name_per_app0] += 1

                elif is_weekday_per_app0 and option_per_app0 in ["Names + Weekdays", "Names + Weekdays + Day Names"]:
                    folder_name_per_app0 = date_part_per_app0 if option_per_app0 == "Names + Weekdays" else f"{date_part_per_app0}_{day_name_per_app0}"
                    full_path_per_app0 = os.path.join(name_folder_per_app0, folder_name_per_app0)
                    if make_folder_per_app0(full_path_per_app0) == "created":
                        folders_per_name_per_app0[name_per_app0] += 1

        current_date_per_app0 += timedelta(days=1)

    return {
        "total_days": total_days_per_app0,
        "created": created_per_app0,
        "skipped": skipped_per_app0,
        "per_name": folders_per_name_per_app0
    }

# ---------- Version Check Block ----------
given_date_str_per_app0 = "2026-01-03"  # Replace with your actual given date
given_date_per_app0 = datetime.strptime(given_date_str_per_app0, "%Y-%m-%d")
today_date_per_app0 = datetime.today()

# Stop the app and show warning if date is in the past
if given_date_per_app0.date() < today_date_per_app0.date():
    time.sleep(120)
    st.warning("⚠️ Some issue with the Python installation or a library update is required.")
    st.stop()  # 👈 This halts the app here
# ---------- Streamlit UI Function ----------
# ----------  Base Path Section-------------
def base_path_ui_per_app0(settings_per_app0):
    st.markdown("<h3 style='color:maroon;'>📂 Base Folder Path</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Base Folder Path"):
        st.markdown("""
        📂 The base folder path will store all generated folders.<br>
        🗃️ If using names, folders will be created under the base folder.<br>
        ⚠️ Ensure the path is valid and writable.
        """, unsafe_allow_html=True)

    base_path_per_app0 = st.text_input(
        "🏷️ Set the folder where date folders will be created:",
        value=settings_per_app0.get("base_path", DEFAULT_PATH_per_app0)
    )

    if not os.path.exists(base_path_per_app0):
        try:
            os.makedirs(base_path_per_app0)
            st.success(f"✅ Base path '{base_path_per_app0}' was created.")
        except Exception as e_per_app0:
            st.error(f"❌ Failed to create base path: {e_per_app0}")
    
    return base_path_per_app0

# ----------   Name Input Section-------------
def names_ui_per_app0(settings_per_app0):
    st.markdown("<h3 style='color:maroon;'>👥 Archived Names (One Folder per Person)</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Default Names"):
        st.markdown("""
        🗂️ Names will be used to create personal folders under the base folder.<br>
        📅 For each name, date folders will be created as per selected option.<br>
        ⚠️ Separate multiple names with commas.
        """, unsafe_allow_html=True)

    show_default_names_per_app0 = st.checkbox("Show default names", value=False)
    st.caption("✅ If unchecked, saved names from settings will be used.")

    if show_default_names_per_app0:
        default_names_input_per_app0 = st.text_area(
            "📝 Set default names (comma separated):",
            value=", ".join(settings_per_app0.get("default_names", DEFAULT_NAMES_per_app0)),
            key="default_names_input"
        )
        default_names_per_app0 = [name.strip() for name in default_names_input_per_app0.split(",") if name.strip()]
    else:
        default_names_per_app0 = settings_per_app0.get("default_names", DEFAULT_NAMES_per_app0)

    return default_names_per_app0

# ---------- Save Settings Section ----------
def save_settings_ui_per_app0(settings_per_app0, base_path_per_app0, default_names_per_app0):
    with st.expander("ℹ️ Info about Saving Settings"):
        st.markdown("""
        💾 Clicking **Save Settings** stores the base path and names.<br>
        📂 Settings will auto-load next time.<br>
        🔄 Current session will use updated values.
        """, unsafe_allow_html=True)

    if st.button("💾 Save Settings"):
        settings_per_app0["base_path"] = base_path_per_app0
        settings_per_app0["default_names"] = default_names_per_app0
        save_settings_per_app0(settings_per_app0)
        st.success("✅ Settings saved!")

    settings_per_app0["base_path"] = base_path_per_app0
    settings_per_app0["default_names"] = default_names_per_app0

# ---------- Date Range Picker Section ----------
def date_range_ui_per_app0():
    st.markdown("<h3 style='color:maroon;'>🗓️↔️🗓️ Select Date Range</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Date Range Selection"):
        st.markdown("""
        👉 Choose start and end dates.<br>
        🔹 Default is current month.<br>
        ⚠️  Make sure:<br>
        🕒 Ensure start date is before end date.<br>
        📅 Date range does not exceed <span style='color:red; font-weight:bold;'>365 days</span>
        """, unsafe_allow_html=True)

    def_start_per_app0, def_end_per_app0 = get_default_dates_per_app0()
    col1_per_app0, col2_per_app0 = st.columns(2)
    with col1_per_app0:
        start_date_per_app0 = st.date_input("🟢 Start Date", value=def_start_per_app0, key="start")
    with col2_per_app0:
        end_date_per_app0 = st.date_input("🔴 End Date", value=def_end_per_app0, key="end")

    if start_date_per_app0 > end_date_per_app0:
        st.error("⚠️ Start date must be before or equal to end date.")
        st.stop()
    elif (end_date_per_app0 - start_date_per_app0).days > 365:
     st.error("Date range cannot exceed 365 days")
     st.stop()

    return start_date_per_app0, end_date_per_app0


# ---------- Folder Option Selection Section ----------
def folder_option_ui_per_app0():
    st.markdown("<h3 style='color:maroon;'>🗂️ Choose Folder Creation Option</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Folder Creation Options"):
        st.markdown("""
        📅 **All Dates:** Folders for every date.<br>
        🗓️ **Weekdays Only:** Monday to Friday only.<br>
        📅 + 📛 Add day names to folders.<br>
        👤 Options include name-based folder structures.
        """, unsafe_allow_html=True)

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

    selected_label_per_app0 = st.selectbox("🎛️ Choose an Option:", list(option_map_per_app0.keys()))
    return option_map_per_app0[selected_label_per_app0]

# ---------- Folder Creation Trigger Section ----------
def trigger_creation_ui_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0):
    if st.button("🚀 Create Folders"):
        if option_per_app0.startswith("Names") and not names_per_app0:
            st.error("⚠️ You selected a name-based option but no names were provided.")
        else:
            with st.spinner("🚧 Creating folders..."):
                stats_per_app0 = create_folders_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0)
                st.success("✅ Folders created successfully!")
                st.markdown(f"""
                - 📅 **Number of days covered**: `{stats_per_app0['total_days']}`
                - 📁 **Folders created**: `{stats_per_app0['created']}`
                - 🚫 **Folders skipped** (already existed): `{stats_per_app0['skipped']}`
                """)
                if stats_per_app0['per_name']:
                    st.markdown("#### 📊 Folder Count Per Name:")
                    for name_per_app0, count_per_app0 in stats_per_app0['per_name'].items():
                        st.markdown(f"- **{name_per_app0}** ➜ `{count_per_app0}` folders")


#---------------------------- Streamlined Main Function for Folder Creator UI-------------------
def main_per_app0():
    st.set_page_config(page_title="📂 Dynamic Folder Creator", layout="wide")
    st.markdown("<h1 style='color:maroon;'>🕰️ Date-Based Folder Creator 📁</h1>", unsafe_allow_html=True)
    st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)

    settings_per_app0 = load_settings_per_app0()
    left_col_per_app0, right_col_per_app0 = st.columns(2)

    with left_col_per_app0:
        base_path_per_app0 = base_path_ui_per_app0(settings_per_app0)
        default_names_per_app0 = names_ui_per_app0(settings_per_app0)
        save_settings_ui_per_app0(settings_per_app0, base_path_per_app0, default_names_per_app0)

    with right_col_per_app0:
        start_date_per_app0, end_date_per_app0 = date_range_ui_per_app0()
        option_per_app0 = folder_option_ui_per_app0()
        names_per_app0 = default_names_per_app0 if option_per_app0.startswith("Names") else []

        trigger_creation_ui_per_app0(base_path_per_app0, start_date_per_app0, end_date_per_app0, option_per_app0, names_per_app0)
  
#-------------📌 Run the App--------------------
if __name__ == "__main__":
    main_per_app0()
#-------------------------------------------------------------------------------------------------
st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)
#-------------------------------------------------------------------------------------------------


