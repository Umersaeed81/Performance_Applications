# 📂 Date-Based Folder Creator using Streamlit

A powerful and interactive Streamlit web app that dynamically creates folders based on date ranges, weekdays, day names, and personalized names. Ideal for managing attendance records, daily logs, or task tracking folders in a systematic and automated way.

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
<pre> git clone https://github.com/your-username/date-folder-creator.git 
cd date-folder-creator  </pre>
