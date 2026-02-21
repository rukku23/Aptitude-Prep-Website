#!/usr/bin/env python3
import os
import time
import argparse
import pandas as pd
import smtplib
import ssl                           # <- IMPORTANT: import ssl
from email.message import EmailMessage

# CONFIG (use env vars if possible)
SENDER_EMAIL = os.getenv("EMAIL_USER") or "siyaramlakshman87@gmail.com"
APP_PASSWORD = os.getenv("EMAIL_PASS") or "zglvuntqsptpowxn"
SMTP_SERVER = os.getenv("SMTP_HOST") or "smtp.gmail.com"
SMTP_PORT   = int(os.getenv("SMTP_PORT") or 465)

RESULT_FILE = "Result.xlsx"   # both files in same folder
DELAY_SECONDS = 0.8

# helpers
def find_col(df, names):
    for n in names:
        for c in df.columns:
            if c.strip().lower() == n.strip().lower():
                return c
    raise KeyError(f"None of {names} found. Columns: {list(df.columns)}")

def build_subject(status):
    s = str(status).strip().lower()
    return "Your Engineering Aptitude Test Result — Accepted" if s.startswith("acc") else "Your Engineering Aptitude Test Result — Update"

def build_body(name, roll, score, status, date=None):
    date_line = f"Date: {date}\n" if date else ""
    return (f"Hi {name},\n\n"
            f"This is regarding your Engineering Aptitude Test.\n\n"
            f"Roll No: {roll}\n"
            f"Score  : {score}\n"
            f"{date_line}"
            f"Status : {status}\n\n"
            "Thank you for taking the test.\n\nRegards,\nPlacement / Admin Team\n")

def send_mail(to_addr, subject, body, dry_run=False):
    if dry_run:
        print(f"[DRY RUN] Would send to {to_addr} | {subject}")
        return True
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        # print full exception for debugging (authentication/connection problems)
        print("SMTP error:", repr(e))
        return False

def main(result_file, dry_run=False):
    if not os.path.exists(result_file):
        print("Result file not found:", result_file); return 1

    df = pd.read_excel(result_file, engine="openpyxl")
    name_col  = find_col(df, ["Name", "name"])
    roll_col  = find_col(df, ["Roll No","Roll","USN","usn"])
    email_col = find_col(df, ["Email","email","E-mail"])
    score_col = find_col(df, ["Score","score"])
    status_col= find_col(df, ["Status","status"])
    # optional date
    date_col = None
    for c in df.columns:
        if c.strip().lower() == "date":
            date_col = c
            break

    print("Loaded:", result_file)
    print("Columns found ->", name_col, roll_col, email_col, score_col, status_col, date_col)

    for idx, row in df.iterrows():
        raw_email = row.get(email_col)
        # skip empty/NaN emails
        if pd.isna(raw_email) or not str(raw_email).strip():
            print(f"[{idx}] SKIP missing email -> {raw_email}")
            continue
        to_addr = str(raw_email).strip()
        if "@" not in to_addr:
            print(f"[{idx}] SKIP invalid email -> {to_addr}")
            continue

        name = row.get(name_col, "")
        roll = row.get(roll_col, "")
        score= row.get(score_col, "")
        status = row.get(status_col, "")
        date = row.get(date_col, "") if date_col else None

        subj = build_subject(status)
        body = build_body(name, roll, score, status, date)

        ok = send_mail(to_addr, subj, body, dry_run=dry_run)
        if ok:
            print(f"[{idx}] Sent -> {to_addr} | {name} | {score} | {status}")
        else:
            print(f"[{idx}] ❌ Failed to send to {to_addr}")

        time.sleep(DELAY_SECONDS)
    print("Done.")
    return 0

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input","-i", help="Path to Result.xlsx", default=RESULT_FILE)
    parser.add_argument("--dry-run", action="store_true", help="Do not actually send")
    args = parser.parse_args()
    main(args.input, dry_run=args.dry_run)
