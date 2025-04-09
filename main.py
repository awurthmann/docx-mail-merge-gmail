#!/usr/bin/env python3

import csv
import smtplib
from email.mime.text import MIMEText
from docx import Document
from docx.oxml.ns import qn
from string import Template
import config
from docx.oxml import OxmlElement


def get_paragraph_html(paragraph):
    html = ""
    rels = paragraph.part.rels
    p_xml = paragraph._p

    for child in p_xml:
        tag = child.tag.split('}')[-1]

        if tag == "hyperlink":
            r_id = child.get(qn('r:id'))
            url = rels[r_id].target_ref if r_id in rels else ""
            link_text = ""
            for r in child.findall(".//w:t", namespaces=paragraph.part.element.nsmap):
                link_text += r.text or ""
            if url:
                html += f'<a href="{url}">{link_text}</a>'

        elif tag == "r":
            text = ""
            bold = italic = underline = False
            color = None
            size = None

            r_fonts = child.find("w:rPr", namespaces=paragraph.part.element.nsmap)
            if r_fonts is not None:
                bold = r_fonts.find("w:b", namespaces=paragraph.part.element.nsmap) is not None
                italic = r_fonts.find("w:i", namespaces=paragraph.part.element.nsmap) is not None
                underline = r_fonts.find("w:u", namespaces=paragraph.part.element.nsmap) is not None

                color_el = r_fonts.find("w:color", namespaces=paragraph.part.element.nsmap)
                if color_el is not None and color_el.get("w:val"):
                    color = color_el.get("w:val")

                size_el = r_fonts.find("w:sz", namespaces=paragraph.part.element.nsmap)
                if size_el is not None and size_el.get("w:val"):
                    try:
                        size = int(size_el.get("w:val")) / 2  # half-points to pt
                    except ValueError:
                        pass

            for t in child.findall(".//w:t", namespaces=paragraph.part.element.nsmap):
                text += t.text or ""

            if bold:
                text = f"<strong>{text}</strong>"
            if italic:
                text = f"<em>{text}</em>"
            if underline:
                text = f"<u>{text}</u>"

            style = ""
            if color:
                style += f"color:#{color};"
            if size:
                style += f"font-size:{int(size)}pt;"
            if style:
                text = f'<span style="{style}">{text}</span>'

            html += text

    # This return must be at base indentation
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
