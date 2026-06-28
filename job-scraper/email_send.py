"""Send outreach emails through the user's Gmail (SMTP + App Password) with a .docx attachment."""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

DOCX_MIME = "vnd.openxmlformats-officedocument.wordprocessingml.document"


def gmail_creds():
    addr = (os.environ.get("GMAIL_ADDRESS") or "").strip()
    pw = os.environ.get("GMAIL_APP_PASSWORD") or ""
    # Google shows App Passwords as "abcd efgh ijkl mnop" — SMTP wants them spaceless.
    pw = "".join(pw.split())
    if not addr or not pw or pw.startswith("xxxx"):
        raise RuntimeError(
            "GMAIL_ADDRESS / GMAIL_APP_PASSWORD not set in .env. Enable 2FA on the Gmail "
            "account and create an App Password (https://myaccount.google.com/apppasswords).")
    return addr, pw


def send_gmail(to, subject, body, attachment_path=None, attachment_name=None):
    """Send a plain-text email via Gmail with an optional .docx attachment."""
    addr, pw = gmail_creds()
    msg = MIMEMultipart()
    msg["From"] = addr
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=DOCX_MIME)
        part.add_header("Content-Disposition", "attachment",
                        filename=attachment_name or os.path.basename(attachment_path))
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(addr, pw)
        s.sendmail(addr, [to], msg.as_string())
