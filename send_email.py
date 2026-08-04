"""
Sends the day's Chartink screener Excel file as an email attachment via Gmail SMTP.

Required environment variables (set as GitHub Actions secrets):
  GMAIL_ADDRESS    - the Gmail address you're sending FROM (e.g. yourname@gmail.com)
  GMAIL_APP_PASSWORD - a 16-character Gmail App Password (NOT your normal password)
  TO_EMAIL         - the address to send the report TO (can be the same as GMAIL_ADDRESS)
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

if not all([GMAIL_ADDRESS, GMAIL_APP_PASSWORD, TO_EMAIL]):
    print("❌  Missing one or more required env vars: GMAIL_ADDRESS, GMAIL_APP_PASSWORD, TO_EMAIL")
    sys.exit(1)

# Find today's generated Excel file
today_str = datetime.date.today().strftime("%Y-%m-%d")
pattern   = f"Chartink_Screener_{today_str}.xlsx"
matches   = glob.glob(pattern) + glob.glob(os.path.join("**", pattern), recursive=True)

if not matches:
    print(f"❌  No output file found matching {pattern} — did the screener script run and save successfully?")
    sys.exit(1)

filepath = matches[0]
filename = os.path.basename(filepath)
print(f"📎  Attaching: {filepath}")

msg = EmailMessage()
msg["Subject"] = f"Chartink Screener Results — {today_str}"
msg["From"]    = GMAIL_ADDRESS
msg["To"]      = TO_EMAIL
msg.set_content(
    f"Hi,\n\nAttached are today's Chartink screener results ({today_str}).\n\n"
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
