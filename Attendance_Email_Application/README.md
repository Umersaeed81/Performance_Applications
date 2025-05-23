# 📥 Outlook Email Saver App
An automation tool built with Python and Streamlit to extract and save emails from Outlook or PST files. Useful for project-based organizations to verify employee engagement, archive client communications, or prepare audit reports.

# 🔍 Features
Select between current Outlook profile or external .pst file

- Filter emails by:
  - Sender address
  - Date range
  - Specific folders (Inbox, Sent Items, etc.)
- Save filtered emails in .msg format to a specified folder
- Automatically resolve the actual SMTP address even for Exchange senders
- Streamlit UI for ease of use

 # 🎯 Use Cases

- Validating employee assignment in temporary or project-based roles
- Backing up correspondence with clients or vendors
- Creating archives for audit preparation
- Extracting task-specific email histories

# 📸 Interface Preview


# ⚙️ Requirements

- Python 3.8 or higher
- Microsoft Outlook installed and configured
- Required Python packages

# 🛠️ Installation

## 1. Clone the repository:
```python
git clone https://github.com/yourusername/outlook-email-saver.git
cd outlook-email-saver
```

## 2. Create virtual environment (optional but recommended):
```python
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

## 3. Install dependencies:
```python
pip install -r requirements.txt
```

# 🚀 How to Run
1. Open Outlook (required for accessing profiles or PSTs)
2. Run the Streamlit app:
```python
streamlit run app.py
```
3. Use the interface to load a PST or Outlook folder, apply filters, and export emails.

# 📁 Output
- Saved emails are exported as `.msg` files
- File names are system-generated and maintain original email metadata
- Output directory is user-defined at runtime

# 💡 Future Plans
- Support for shared mailboxes
- Export as PDF or text summaries
- Custom file naming options
- Bulk sender filters and keyword search

# 📌 Limitations
- Only works on systems where Outlook is installed
- Not tested with .ost files
- Requires proper permissions for Outlook folders

# 🙋 Contributing
We welcome contributions! Please fork the repo and create a pull request or open an issue for suggestions and bugs.

# 🔗 Contact
Created by [Umer Saeed]

📧 Email: umersaeed81@hotmail.com

🌐 LinkedIn: https://www.linkedin.com/in/engumersaeed/
