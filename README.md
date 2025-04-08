# DOCX Mail Merge with Gmail (Python)

![Python](https://img.shields.io/badge/python-3.9-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey)

Send personalized rich-text emails using Gmail, a .docx template, and a CSV file — all in Python.

**Project URL:** https://github.com/awurthmann/docx-mail-merge-gmail

## ✨ Features
- Use a `.docx` file as your email template — created in **Microsoft Word or Google Docs** (just download as `.docx`)
- Personalize with variables like `$firstname`, `$lastname`, `$company`, `$title`
- Load recipient list from a CSV file
- Send emails securely via Gmail SMTP with an App Password
- Preserve links from your Word document

## 💡 Setup Instructions

### 1. Clone the Repository
```bash
git clone https://github.com/awurthmann/docx-mail-merge-gmail.git
cd docx-mail-merge-gmail
```

### 2. Set Up App Password (Gmail)
1. Go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
2. Make sure 2-Step Verification is enabled on your Google account
3. Choose **App: Mail** and **Device: Other (e.g., MyComputer)**
4. Copy the 16-character app password

### 3. Create a `config.py` File
1. Copy the example config:
   ```bash
   cp example_config.py config.py
   ```
2. Edit `config.py` and fill in your info:
```python
YOUR_EMAIL = 'your-email@gmail.com'
APP_PASSWORD = 'your-app-password'
EMAIL_SUBJECT = 'Your email subject here'
DOCX_PATH = 'email_template.docx'
CSV_PATH = 'recipients.csv'
```

> ⚠️ `config.py` is included in `.gitignore` to keep sensitive info safe.

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

### 5. Run the Script
```bash
python send_mail_merge_smtp.py
```

## 💬 Optional: Change the Sender Name in Gmail
If your name isn’t displayed correctly in the outgoing emails sent through the Gmail integration, you can fix it easily:

1. Log in to your Gmail account.
2. Click on the **gear icon**, then **See all settings**.
3. Navigate to the **Accounts** tab.
4. Find the **Send mail as** section and click **edit info**.
5. In the popup, enter your preferred name in the empty box.
6. Select the box next to your preferred name and click **Save Changes**.

## 📄 Example

**recipients.csv**:
```csv
email,firstname,lastname,company,title
alice@example.com,Alice,Johnson,TechCo,CTO
bob@example.com,Bob,Smith,SecureCorp,Engineer
```

**email_template.docx**:
```
Hi $firstname $lastname,

It was great meeting you at $company. Hope you enjoyed our talk about modern security practices as a $title.
Click here to visit our site: https://example.com

Best,
Aaron Wurthmann
```

## ✅ Output
Sends a personalized email to each recipient using Gmail's SMTP service, and preserves rich-text formatting including hyperlinks.

---

MIT License

---

Feel free to contribute, fork, or adapt to your needs!