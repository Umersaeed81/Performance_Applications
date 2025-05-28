![Python Version](https://img.shields.io/badge/python-3.9+-blue)
![License](https://img.shields.io/badge/license-MIT-green)
---------------------------------
# 📬 Building the Ultimate Outlook Email Filter & Downloader with Streamlit

## 📖 The Story Behind the App

Every corporate professional has experienced the overwhelming flood of emails—newsletters, CCs, missed attachments, follow-up requests, and, of course, those critical action items buried under dozens of threads.

Imagine being handed a powerful yet intuitive app that extracts only the emails you care about, with pinpoint precision.

This is the story behind the creation of our **Streamlit-powered Outlook Email Filter and Downloader**.

What started as a small utility to download emails from specific folders quickly evolved into something much more powerful. The app learned to respect your filters—whether it be:

- Sender addresses  
- Unread/read status  
- Email size  
- Presence and content of attachments  
- Importance levels and follow-up flags  

Soon, it wasn't just an app. It became your personal assistant for email mining.

---

## 🔙 May 23, 2025 — The Minimalist Start

On May 23, 2025, we released the first version of our Outlook Streamlit app. It was a smart but minimalistic tool, introduced in a [LinkedIn Article](https://www.linkedin.com/pulse/automating-outlook-email-extraction-project-based-umer-saeed-prphf/?trackingId=E87bQYe%2F9qNJ22MMoB5s7A%3D%3D).

### 💡 Version 1 Highlights:
- ✅ Accepted a list of sender email addresses  
- 📅 Grouped emails per sender, per date  
- 📉 Downloaded only the smallest-sized email in each group  

This early version solved real-world clutter in crowded inboxes. It worked well for batch operations and targeted email archiving.

But users wanted more control. More flexibility. And that’s where the journey truly began.

---

## 🚀 Today’s Upgrade — A New Era in Email Filtering

The new version of the app isn’t just about filtering by size. It’s a **dashboard of intelligent controls** for navigating the chaos of corporate email overload.

---

## 🗂️ Folder Selection Panel

### 💼 Use Current PST
Access your default Outlook inbox directly.

### 🧭 Browse PST File
Load any external `.pst` archive via full file path (e.g., `E:\Email\Reports.pst`).

### 🧠 Smart Validation
Ensures only one of the above is selected—warns if both or neither are chosen.

---

## 📂 Folder & Subfolder Picker

- 🔄 Recursive fetching of folders and subfolders  
- ✅ Full path visibility like `Inbox/Clients/Q1 Reports`  
- 🧾 Checkbox UI to select multiple folders  
- 📊 Displays total folders selected for clarity  

---

## 🆔 Sender Filtering

- 👤 Multiple Sender Input: Enter one or more sender email addresses for pinpoint targeting.  
- 🧩 Exact Match Logic: Filters strictly on exact sender matches—no partial lookups.  
- 💾 Smart JSON Memory: Auto-saves frequently used addresses locally and loads them next time.  

---

## 📆 Date Range Filter

- 📅 Custom Date Selection: Select precise start and end dates.  
- 🕓 Based on Received Date: Ensures filtering works off the actual email received timestamp.  
- 🎯 Daily Sender Targeting: Combine with sender filters for per-day, per-sender extraction.  
- 🗓️ Defaults to Today: Speeds up daily report use cases with default date set to current day.  

---

## 📝 Subject Filtering

### 🔤 Match Modes

Choose from:
- 🔍 Contains  
- 🚫🔍 Does Not Contain  
- 🏁 Starts With  
- 🏁➡️ Ends With  
- 🎯 Exact Match  

- ➕ Multi-Keyword Support  
- 🧠 Case Sensitivity Toggle  
- 🛠️ Combine with Sender for powerful targeting  

---

## 📨 Read / Unread Status

Choose to focus only on:
- 📩 Unread emails  
- 📬 Read emails  

🔄 Works in harmony with all other filters.

---

## 🚦 Email Importance Filtering

Filter by priority level:
- 🟢 Low  
- 🟡 Normal  
- 🔴 High  

✅ Useful for focusing on high-priority, time-sensitive communications.

---

## 🚩 Follow-Up Flag Filter

Track actionable emails:
- 🚫 No Flag  
- ✅ Completed  
- 🚩 Flagged  

Perfect for reviewing tasks and follow-ups.

---

## 📏 Email Size Filter

### 📐 Min or Max Mode

Choose to download either:
- 🟢 Minimum-size email per group (efficient)
- 🔴 Maximum-size email per group (full context)

### 🎯 Advanced Comparison Options

Choose from:
- 🟰 Equals  
- 🟰🚫 Not Equals  
- 🔽 Less Than  
- 🔼 Greater Than  
- 🔽🟰 Less Than or Equal To  
- 🔼🟰 Greater Than or Equal To  
- 🔁 Between (range)

### 🔎 Why It Matters

Perfect for isolating:
- Lightweight emails for summaries  
- Heavy emails for attachments or in-depth threads  

---

## 📎 Attachment Filters

### 📥 Presence Control

Choose to:
- ✅ Include only emails with attachments  
- 🚫 Exclude all emails that contain attachments  

### 🧾 File Type Filtering

Restrict to specific types:
- 📄 `.pdf`
- 📊 `.xlsx`
- 📝 `.docx`
- 📦 `.zip`
- ➕ And more…

> 🧠 Tip: Enter multiple extensions separated by commas.

### 🧠 Attachment Name Filtering

Choose match mode:
- 🔍 Contains  
- 🚫🔍 Does Not Contain  
- 🏁 Starts With  
- 🏁➡️ Ends With  
- 🎯 Exact Match  

📌 Combine with sender or subject filters for ultra-specific targeting.

---

## 💾 Download Settings

### 📁 Set Your Path

Choose where emails/attachments are saved.  
🛠️ If it doesn't exist, the app creates the folder automatically.

### 🔄 Download Modes

- 🆕 Append: Add only new emails  
- ♻️ Overwrite: Replace existing files with fresh ones  

💡 Great for daily automation or cleanup.

---

## ⏬ Execution + Missing Summary

Click `⬇️ Download Emails`, and the app:
- 🚀 Applies all filters  
- ⏱️ Tracks time taken  
- ✅ Shows total downloaded  
- ❌ Lists missing emails (grouped by sender/date)  

---

## 🎉 The Result

With this app, you're not just filtering emails—you’re **taking back control** of your Outlook workflow.

Ideal for:
- ✅ Analysts with daily reports  
- ✅ Managers tracking follow-ups  
- ✅ Auditors reviewing email trails  
- ✅ Anyone needing less noise and more signal  

---

## 🔮 What’s Next: Ideas for Future Enhancements

- 📦 Batch Export to ZIP  
- 🔔 Email Alerts & Scheduling  
- 🕵️‍♂️ Attachment Content Search  
- 📊 Analytics Dashboard  
- 🌐 Web App via FastAPI or Streamlit Cloud  

---

## ✅ Conclusion: A Smart Companion for the Modern Inbox

In the fast-paced world of corporate communication, finding the right email shouldn’t feel like hunting for a needle in a haystack.

This app transforms that struggle into a **streamlined, intelligent experience**—offering precision via Outlook + Streamlit.

🔧 Whether you're an analyst, manager, or compliance officer—this tool is your new digital assistant.

> The inbox isn’t getting smaller, but with the right tools, it can definitely get smarter. 💼📥

---

## 🧰 Tech Stack & Connect

🔗 **Built With**:  
🐍 Python · 🌐 Streamlit · 📬 win32com.client (Outlook Automation)

🧠 **Designed for modularity**:  
Scalable and maintainable—ready for enterprise-level deployment and continuous evolution.

💬 **Got feedback, ideas, or questions?**  
Let’s connect! I’m always open to collaboration and suggestions.

📦 **Explore the Code on GitHub**:  
👉 [GitHub Repository – Outlook Email Filter & Downloader](https://github.com/Umersaeed81/Performance_Applications/blob/main/Attendance_Email_Application/Attendance_Email_Application_05.py)

---

> ✨ *Star the repo if you find it useful, and feel free to fork or contribute!*
