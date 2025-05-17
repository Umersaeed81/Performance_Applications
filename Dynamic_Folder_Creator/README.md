# 📂 Date-Based Folder Creator using Streamlit

A powerful and interactive Streamlit web app that dynamically creates folders based on date ranges, weekdays, day names, and personalized names. Ideal for managing attendance records, daily logs, or task tracking folders in a systematic and automated way.

[Application Code](https://github.com/Umersaeed81/Performance_Applications/blob/main/Dynamic_Folder_Creator/Folder_Creation_Application.py)

## 🚀 Features

- 📆 Date Range Selection — Choose any start and end date to generate folders within that range.
- 🗓️ Weekday Filtering — Generate folders only for weekdays (Monday–Friday).
- 📛 Day Name Tagging — Append day names (e.g., Monday, Tuesday) to folder names for better context.
- 👤 Name-Based Folder Structure — Create folders for individuals (e.g., team members, students) with nested date-wise folders inside.
- 💾 Settings Persistence — Save base path and default names to reuse automatically in future sessions.

## 📁 Folder Creation Options

| Option Label                               | Description                                                  |
| ------------------------------------------ | ------------------------------------------------------------ |
| 📅 All Dates                               | Folder for each date in the selected range                   |
| 🗓️ Weekdays Only                          | Folder for weekdays (Mon–Fri) only                           |
| 📅 + 📛 All Dates + Day Names              | Folder for all dates with appended day names                 |
| 🗓️ + 📛 Weekdays Only + Day Names         | Weekdays only with appended day names                        |
| 👤 + 📅 Names + All Dates                  | Creates personal folders, each containing date-based folders |
| 👤 + 🗓️ Names + Weekdays                  | Name-based folders with weekday-only subfolders              |
| 👤 + 📅 + 📛 Names + All Dates + Day Names | Name-based folders with all dates and day names              |
| 👤 + 🗓️ + 📛 Names + Weekdays + Day Names | Name-based folders for weekdays with day names               |

## 🛠️ Tech Stack

- Python 3.9+
- Streamlit – for UI
- os, json, datetime, calendar – standard Python libraries for folder operations and date handling

## 🔧 How to Use

### 1. Clone this Repository:

```python
git clone https://github.com/your-username/date-folder-creator.git
cd date-folder-creator
```

### 2. Install Dependencies:

```python
pip install streamlit
```
### 3. Run the App:

```python
streamlit run app.py
```
### 4. Access the Web UI
Open the given localhost URL in your browser (usually `http://localhost:8501/`).

## 📂 Example Output
If you choose:
- Names: Umer, Ali
- Date Range: 2025-05-01 to 2025-05-05
- Option: 👤 + 📅 + 📛 Names + All Dates + Day Names

It will create folders like:

```python
E:/Attendance/Umer/2025-05-01_Thursday
E:/Attendance/Umer/2025-05-02_Friday
...
E:/Attendance/Ali/2025-05-01_Thursday
...
```

## 🧠 Tips

✅ Use the Save Settings button to store frequently used paths and names.
🧼 Make sure paths are valid and the script has permission to write to the chosen directory.


![Python Version](https://img.shields.io/badge/python-3.9+-blue)
![License](https://img.shields.io/badge/license-MIT-green)

