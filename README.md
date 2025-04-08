# docx-mail-merge-gmail
Send personalized rich-text emails using Gmail, Microsoft Word (.docx) templates, and CSV data — all with a simple Python script.

## ✨ Features
- Use Microsoft Word (.docx) to author your email template
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

### 3. Configure Your Settings
Edit `config.py` and set:
```python
YOUR_EMAIL = 'your-email@gmail.com'
APP_PASSWORD = 'your-app-password'
EMAIL_SUBJECT = 'Your email subject here'
DOCX_PATH = 'email_template.docx'
CSV_PATH = 'recipients.csv'
```
We are using a config.py vs an .env file here for "reasons" so be sure to keep this password safe and only run this script on a secured system. In fact, you should consider disabling or chaning the app password after each use. 

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

### 5. Run the Script
```bash
python send_mail_merge_smtp.py
```

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