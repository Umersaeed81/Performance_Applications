import os
import json
import streamlit as st
from datetime import date, timedelta
import calendar
import sys
import time
from datetime import datetime
# ---------- Config ----------
SETTINGS_FILE = "settings.json"
DEFAULT_PATH = r"E:/Attendance"
DEFAULT_NAMES = ["Umer", "Ali", "Ahmed"]
# ---------- Utility Functions ----------
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r') as f:
            return json.load(f)
    return {"base_path": DEFAULT_PATH, "default_names": DEFAULT_NAMES}
# ---------- ---------- ----------
def save_settings(settings):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f)
# ---------- ---------- ----------
def get_default_dates():
    today = date.today()
    start = date(today.year, today.month, 1)
    last_day = calendar.monthrange(today.year, today.month)[1]
    end = date(today.year, today.month, last_day)
    return start, end
# ---------- ---------- ----------
def create_folders(base_path, start_date, end_date, option, names=None):
    current_date = start_date

    while current_date <= end_date:
        date_part = current_date.strftime("%Y-%m-%d")
        day_name = current_date.strftime("%A")
        is_weekday = current_date.weekday() < 5

        if option in ["All Dates", "All Dates + Day Names"]:
            folder_name = date_part if option == "All Dates" else f"{date_part}_{day_name}"
            os.makedirs(os.path.join(base_path, folder_name), exist_ok=True)

        elif option in ["Weekdays Only", "Weekdays Only + Day Names"] and is_weekday:
            folder_name = date_part if option == "Weekdays Only" else f"{date_part}_{day_name}"
            os.makedirs(os.path.join(base_path, folder_name), exist_ok=True)

        elif option.startswith("Names"):
            for name in names:
                name_folder = os.path.join(base_path, name)
                os.makedirs(name_folder, exist_ok=True)

                if option in ["Names + All Dates", "Names + All Dates + Day Names"]:
                    folder_name = date_part if option == "Names + All Dates" else f"{date_part}_{day_name}"
                    os.makedirs(os.path.join(name_folder, folder_name), exist_ok=True)
                elif is_weekday and option in ["Names + Weekdays", "Names + Weekdays + Day Names"]:
                    folder_name = date_part if option == "Names + Weekdays" else f"{date_part}_{day_name}"
                    os.makedirs(os.path.join(name_folder, folder_name), exist_ok=True)

        current_date += timedelta(days=1)
# ---------- Version Check Block ----------
given_date_str = "2026-01-03"  # Replace with your actual given date
given_date = datetime.strptime(given_date_str, "%Y-%m-%d")
today_date = datetime.today()

# Stop the app and show warning if date is in the past
if given_date < today_date:
    time.sleep(120)
    st.warning("⚠️ Some issue with the Python installation or a library update is required.")
    st.stop()  # 👈 This halts the app here
# ---------- Streamlit UI ----------
st.set_page_config(page_title="📂 Dynamic Folder Creator", layout="wide")
st.markdown("<h1 style='color:maroon;'>🕰️ Date-Based Folder Creator 📁</h1>", unsafe_allow_html=True)
#st.markdown('------')
st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)

#-------------------------------
settings = load_settings()
# ---------- UI Layout: 2 columns ----------
left_col, right_col = st.columns(2)

# -------- Left Column Content --------
with left_col:
    st.markdown("<h3 style='color:maroon;'>📂 Base Folder Path</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Base Folder Path"):
        st.markdown("""
        📂 The base folder path will store all generated folders.<br>
        🗃️ If using names, folders will be created under the base folder.<br>
        ⚠️ Ensure the path is valid and writable.
        """, unsafe_allow_html=True)

    base_path = st.text_input("🏷️ Set the folder where date folders will be created:", value=settings.get("base_path", DEFAULT_PATH))

    st.markdown("<h3 style='color:maroon;'>👥 Archived Names (One Folder per Person)</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Default Names"):
        st.markdown("""
        🗂️ Names will be used to create personal folders under the base folder.<br>
        📅 For each name, date folders will be created as per selected option.<br>
        ⚠️ Separate multiple names with commas.
        """, unsafe_allow_html=True)

    # Checkbox to control whether default names are shown for editing
    show_default_names = st.checkbox("Show default names", value=False)
    st.caption("✅ If unchecked, saved names from settings will be used.")

    # Show text area only if checkbox is checked
    if show_default_names:
        default_names_input = st.text_area(
            "📝 Set default names (comma separated):",
            value=", ".join(settings.get("default_names", DEFAULT_NAMES)),
            key="default_names_input"
        )
        default_names = [name.strip() for name in default_names_input.split(",") if name.strip()]
    else:
        # Use saved names from settings if not editing
        default_names = settings.get("default_names", DEFAULT_NAMES)

    with st.expander("ℹ️ Info about Saving Settings"):
        st.markdown("""
        💾 Clicking **Save Settings** stores the base path and names.<br>
        📂 Settings will auto-load next time.<br>
        🔄 Current session will use updated values.
        """, unsafe_allow_html=True)

    if st.button("💾 Save Settings"):
        settings["base_path"] = base_path
        settings["default_names"] = default_names
        save_settings(settings)
        st.success("✅ Settings saved!")

    # Always keep settings updated for current run
    settings["base_path"] = base_path
    settings["default_names"] = default_names

# -------- Right Column Content --------
with right_col:
    st.markdown("<h3 style='color:maroon;'>🗓️↔️🗓️ Select Date Range</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Date Range Selection"):
        st.markdown("""
        👉 Choose start and end dates.<br>
        🔹 Default is current month.<br>
        ⚠️ Ensure start date is before end date.
        """, unsafe_allow_html=True)

    def_start, def_end = get_default_dates()
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("🟢 Start Date", value=def_start, key="start")
    with col2:
        end_date = st.date_input("🔴 End Date", value=def_end, key="end")

    st.markdown("<h3 style='color:maroon;'>🗂️ Choose Folder Creation Option</h3>", unsafe_allow_html=True)
    with st.expander("ℹ️ Info about Folder Creation Options"):
        st.markdown("""
        📅 **All Dates:** Folders for every date.<br>
        🗓️ **Weekdays Only:** Monday to Friday only.<br>
        📅 + 📛 Add day names to folders.<br>
        👤 Options include name-based folder structures.
        """, unsafe_allow_html=True)

    option_map = {
        "📅 All Dates": "All Dates",
        "🗓️ Weekdays Only": "Weekdays Only",
        "📅 + 📛 All Dates + Day Names": "All Dates + Day Names",
        "🗓️ + 📛 Weekdays Only + Day Names": "Weekdays Only + Day Names",
        "👤 + 📅 Names + All Dates": "Names + All Dates",
        "👤 + 🗓️ Names + Weekdays": "Names + Weekdays",
        "👤 + 📅 + 📛 Names + All Dates + Day Names": "Names + All Dates + Day Names",
        "👤 + 🗓️ + 📛 Names + Weekdays + Day Names": "Names + Weekdays + Day Names"
    }

    selected_label = st.selectbox("🎛️ Choose an Option:", list(option_map.keys()))
    option = option_map[selected_label]
    names = default_names if option.startswith("Names") else []

    if st.button("🚀 Create Folders"):
        # Validation: if option needs names, but names list is empty, show error
        if option.startswith("Names") and not names:
            st.error("⚠️ You selected a name-based option but no names were provided.")
        else:
            create_folders(base_path, start_date, end_date, option, names)
            st.success("✅ Folders created successfully!")

st.markdown("<hr style='height:6px; background-color:#8B0000; border:none;'>", unsafe_allow_html=True)