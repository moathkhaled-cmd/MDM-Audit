"""
Email report sender
Reads data/report.xlsx and sends it as attachment to recipients defined in environment variables.
Environment variables required:
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, RECIPIENTS (comma-separated)
Optionally set SENDER (defaults to SMTP_USER)

Usage:
  python3 scripts/email_report.py --file data/report.xlsx --subject "MDM Audit Report"
"""

import os
import argparse
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


def send_email(smtp_host, smtp_port, smtp_user, smtp_pass, sender, recipients, subject, body, attachment_path):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg.set_content(body)

    # attach file
    with open(attachment_path, 'rb') as f:
        data = f.read()
    maintype = 'application'
    subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=Path(attachment_path).name)

    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', default='data/report.xlsx')
    parser.add_argument('--subject', default='MDM Audit Report')
    parser.add_argument('--body', default='Attached is the latest MDM audit report.')
    args = parser.parse_args()

    smtp_host = os.getenv('SMTP_HOST')
    smtp_port = os.getenv('SMTP_PORT')
    smtp_user = os.getenv('SMTP_USER')
    smtp_pass = os.getenv('SMTP_PASS')
    recipients = os.getenv('RECIPIENTS')
    sender = os.getenv('SENDER') or smtp_user

    if not all([smtp_host, smtp_port, smtp_user, smtp_pass, recipients]):
        print('Please set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS and RECIPIENTS in environment (or .env)')
    else:
        recips = [r.strip() for r in recipients.split(',') if r.strip()]
        send_email(smtp_host, smtp_port, smtp_user, smtp_pass, sender, recips, args.subject, args.body, args.file)
        print('Email sent to', recips)
