"""
Sends the day's screener Excel file as an email attachment via Gmail SMTP.
Shared by both the Chartink and Finviz workflows — which report to send is
controlled by the FILE_PREFIX / REPORT_NAME env vars set in each workflow file.

Required environment variables (set as GitHub Actions secrets):
  GMAIL_ADDRESS       - the Gmail address you're sending FROM (e.g. yourname@gmail.com)
  GMAIL_APP_PASSWORD  - a 16-character Gmail App Password (NOT your normal password)
  TO_EMAIL            - the address to send the report TO (can be the same as GMAIL_ADDRESS)

Optional environment variables (set per-workflow, sensible defaults shown):
  FILE_PREFIX   - filename prefix to look for, default "Chartink_Screener"
  REPORT_NAME   - human-readable name used in the subject/body, default "Chartink Screener"
"""

import os
import sys
import glob
import smtplib
import datetime
from email.message import EmailMessage

GMAIL_ADDRESS      = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
TO_EMAIL            = os.environ.get("TO_EMAIL")
FILE_PREFIX         = os.environ.get("FILE_PREFIX", "Chartink_Screener")
REPORT_NAME         = os.environ.get("REPORT_NAME", "Chartink Screener")

if not all([GMAIL_ADDRESS, GMAIL_APP_PASSWORD, TO_EMAIL]):
    print("❌  Missing one or more required env vars: GMAIL_ADDRESS, GMAIL_APP_PASSWORD, TO_EMAIL")
    sys.exit(1)

# Find today's generated Excel file
today_str = datetime.date.today().strftime("%Y-%m-%d")
pattern   = f"{FILE_PREFIX}_{today_str}.xlsx"
matches   = glob.glob(pattern) + glob.glob(os.path.join("**", pattern), recursive=True)

if not matches:
    print(f"❌  No output file found matching {pattern} — did the screener script run and save successfully?")
    sys.exit(1)

filepath = matches[0]
filename = os.path.basename(filepath)
print(f"📎  Attaching: {filepath}")

msg = EmailMessage()
msg["Subject"] = f"{REPORT_NAME} Results — {today_str}"
msg["From"]    = GMAIL_ADDRESS
msg["To"]      = TO_EMAIL
msg.set_content(
    f"Hi,\n\nAttached are today's {REPORT_NAME} results ({today_str}).\n\n"
    f"This email was sent automatically by your GitHub Actions workflow.\n"
)

with open(filepath, "rb") as f:
    file_data = f.read()

msg.add_attachment(
    file_data,
    maintype="application",
    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    filename=filename,
)

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        smtp.send_message(msg)
    print(f"✅  Email sent to {TO_EMAIL}")
except Exception as e:
    print(f"❌  Failed to send email: {e}")
    sys.exit(1)
