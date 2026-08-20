#!/usr/bin/env python3
"""
send_email_report.py

Build and send an HTML email report summarizing the audit results.
- Reads the audited CSV file from ./run_data/<OUTPUT_CSV_NAME>
- Builds a compact HTML table and highlights rows that contain the phrase
  "No brochure — audited from web" (light yellow background inline style).
- Sends email using SMTP. SMTP credentials are read from environment variables:
  EMAIL_SMTP_HOST, EMAIL_SMTP_PORT, EMAIL_USER, EMAIL_PASS, EMAIL_FROM, EMAIL_FROM_NAME

Usage from CLI (in GitHub Actions):
  python send_email_report.py --attach ./run_data/AUDITED_Arabwheels_MDM_data.csv --to ops@example.com --send

If --send is omitted the script prints the HTML body to stdout.
"""

import os
import argparse
import csv
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

OUTPUT_CSV_NAME = os.environ.get('OUTPUT_CSV_NAME', 'AUDITED_Arabwheels_MDM_data.csv')
DEFAULT_ATTACH = os.path.join('run_data', OUTPUT_CSV_NAME)

SMTP_HOST = os.environ.get('EMAIL_SMTP_HOST')
SMTP_PORT = int(os.environ.get('EMAIL_SMTP_PORT', '587'))
EMAIL_USER = os.environ.get('EMAIL_USER')
EMAIL_PASS = os.environ.get('EMAIL_PASS')
FROM_NAME = os.environ.get('EMAIL_FROM_NAME', 'MDM Audit Bot')
FROM_ADDR = os.environ.get('EMAIL_FROM', EMAIL_USER)


def build_html_table(csv_path, max_rows=500):
    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        headers = reader.fieldnames

    html = ['<table style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:12px;width:100%">']
    # header
    html.append('<thead><tr>')
    for h in headers:
        html.append(f'<th style="border:1px solid #ddd;padding:6px;background:#f2f2f2;text-align:left">{h}</th>')
    html.append('</tr></thead><tbody>')

    for i, r in enumerate(rows):
        if i >= max_rows:
            break
        note = r.get('note', '') or ''
        highlight = 'background:#FFF7B0;' if 'No brochure — audited from web' in note else ''
        html.append(f'<tr style="{highlight}">')
        for h in headers:
            cell = (r.get(h) or '')
            # escape minimal HTML chars
            cell = str(cell).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            html.append(f'<td style="border:1px solid #ddd;padding:6px;vertical-align:top">{cell}</td>')
        html.append('</tr>')

    html.append('</tbody></table>')
    return '\n'.join(html)


def send_email(to_addrs, subject, html_body, attach_path=None):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr((FROM_NAME, FROM_ADDR))
    msg['To'] = ', '.join(to_addrs)
    msg.set_content('Audit report attached. Please open the HTML version for details.')
    msg.add_alternative(html_body, subtype='html')

    if attach_path and os.path.exists(attach_path):
        with open(attach_path, 'rb') as f:
            data = f.read()
        msg.add_attachment(data, maintype='text', subtype='csv', filename=os.path.basename(attach_path))

    if not SMTP_HOST or not EMAIL_USER or not EMAIL_PASS:
        raise RuntimeError('SMTP configuration missing. Set EMAIL_SMTP_HOST, EMAIL_USER, EMAIL_PASS.')

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(EMAIL_USER, EMAIL_PASS)
        s.send_message(msg)


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--attach', default=DEFAULT_ATTACH)
    p.add_argument('--to', required=True)
    p.add_argument('--subject', default='MDM Audit Report')
    p.add_argument('--send', action='store_true')
    args = p.parse_args()

    if not os.path.exists(args.attach):
        raise SystemExit(f'Attachment not found: {args.attach}')

    html_table = build_html_table(args.attach)
    summary = '<p>Attached is the audit CSV and a summary table below. Rows highlighted in light yellow were audited from web (no brochure found).</p>'
    html_body = summary + html_table

    if args.send:
        to_list = [x.strip() for x in args.to.split(',')]
        send_email(to_list, args.subject, html_body, attach_path=args.attach)
        print('Email sent to', to_list)
    else:
        print(html_body)


if __name__ == '__main__':
    main()
