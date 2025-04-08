#!/usr/bin/env python3

import csv
import smtplib
from email.mime.text import MIMEText
from docx import Document
from docx.oxml.ns import qn
from string import Template
import config

def get_paragraph_html(paragraph):
    html = ""
    for run in paragraph.runs:
        text = run.text
        if not text.strip():
            continue

        r = run._element
        hyperlink = r.find(".//w:hyperlink", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if hyperlink is not None:
            r_id = hyperlink.get(qn('r:id'))
            rel = paragraph.part.rels.get(r_id)
            if rel:
                url = rel.target_ref
                text = f'<a href="{url}">{text}</a>'

        # Apply text formatting
        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.underline:
            text = f"<u>{text}</u>"

        # Inline style: color and font size
        style = ""
        if run.font.color and run.font.color.rgb:
            style += f"color:#{run.font.color.rgb};"
        if run.font.size:
            pt_size = int(run.font.size.pt)
            style += f"font-size:{pt_size}px;"
        if style:
            text = f'<span style="{style}">{text}</span>'

        html += text

    # Check if this is a list paragraph
    if paragraph.style.name.lower().startswith("list"):
        return f"<li>{html}</li>"
    else:
        return f"<p>{html}</p>"

def read_docx_template(filepath):
    doc = Document(filepath)
    html_parts = [get_paragraph_html(p) for p in doc.paragraphs]
    return "\n".join(html_parts)

def read_csv(filepath):
    with open(filepath, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def send_email(to_address, subject, html_body):
    msg = MIMEText(html_body, 'html')
    msg['Subject'] = subject
    msg['From'] = config.YOUR_EMAIL
    msg['To'] = to_address

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(config.YOUR_EMAIL, config.APP_PASSWORD.replace(' ', ''))
            server.send_message(msg)
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Error sending to {to_address}: {e}")

def main():
    raw_html = read_docx_template(config.DOCX_PATH)
    template = Template(raw_html)
    recipients = read_csv(config.CSV_PATH)

    for row in recipients:
        personalized_html = template.safe_substitute(row)
        send_email(row['email'], config.EMAIL_SUBJECT, personalized_html)

if __name__ == '__main__':
    main()
