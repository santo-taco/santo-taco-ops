"""
Send a weekly-run summary email via Gmail API (service account + domain delegation).
Called by run_all_weekly.sh after all scripts finish.

Usage:
    python notify.py [--log weekly_run.log]
"""
import argparse
import base64
import os
import sys
from email.mime.text import MIMEText
from pathlib import Path

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

load_dotenv(Path(__file__).parent / ".env")

GOOGLE_CREDS_FILE = os.path.join(os.path.dirname(__file__), "google_credentials.json")
SEND_AS            = "robyn@eatsantotaco.com"
TO                 = "robyn@eatsantotaco.com"
LOG_TAIL_LINES     = 60


def _gmail():
    creds = Credentials.from_service_account_file(
        GOOGLE_CREDS_FILE,
        scopes=["https://www.googleapis.com/auth/gmail.send"],
        subject=SEND_AS,
    )
    return build("gmail", "v1", credentials=creds)


def _tail(path: str, n: int) -> str:
    try:
        lines = Path(path).read_text(errors="replace").splitlines()
        return "\n".join(lines[-n:])
    except Exception:
        return "(log not found)"


def send_summary(log_path: str, failed: list[str]) -> None:
    tail = _tail(log_path, LOG_TAIL_LINES)

    status = "completed with errors" if failed else "completed successfully"
    subject = f"Santo Taco weekly ops — {status}"

    body_lines = [f"Weekly ops run {status}.", ""]
    if failed:
        body_lines += [
            f"⚠ The following scripts failed after {3} attempts and need to be re-run manually:",
            *[f"  • {s}" for s in failed],
            "",
        ]
    body_lines += [
        f"Last {LOG_TAIL_LINES} lines of weekly_run.log:",
        "─" * 60,
        tail,
    ]

    msg = MIMEText("\n".join(body_lines))
    msg["to"]      = TO
    msg["from"]    = SEND_AS
    msg["subject"] = subject

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    _gmail().users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"[notify] Email sent to {TO} — {subject}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--log", default=os.path.join(os.path.dirname(__file__), "weekly_run.log"))
    parser.add_argument("--failed", default="", help="Comma-separated list of scripts that failed after all retries")
    args = parser.parse_args()
    failed = [s.strip() for s in args.failed.split(",") if s.strip()]
    try:
        send_summary(args.log, failed)
    except Exception as exc:
        print(f"[notify] Failed to send email: {exc}", file=sys.stderr)
        sys.exit(1)
