"""
pull_webengage_gmail.py
Fetches WebEngage daily report emails from Gmail, downloads zip files,
extracts CSVs, and writes to Google Sheets tabs.

Reads from: Gmail (agrimmarketing2023@gmail.com)
Writes to:  Webengage_Reports sheet (WEBENGAGE_SHEET_ID)
  - Whatsapp_Campaign
  - Push_Campaign
  - Journey_Campaign
  - Inapp_Campaign
"""

import os
import io
import re
import csv
import json
import zipfile
import base64
import urllib.parse
import requests
import pytz
from datetime import datetime, timedelta
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread
from google.oauth2.service_account import Credentials as ServiceCredentials

load_dotenv()

# ── Constants ──────────────────────────────────────────────────────────────────

IST = pytz.timezone("Asia/Kolkata")

GMAIL_SCOPES           = ["https://www.googleapis.com/auth/gmail.readonly"]
GMAIL_CREDENTIALS_PATH = os.getenv("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json")
GMAIL_TOKEN_PATH       = os.getenv("GMAIL_TOKEN_PATH", "gmail_token.json")
WEBENGAGE_SENDER       = os.getenv("WEBENGAGE_SENDER_EMAIL", "noreply@webengage.com")
WEBENGAGE_SHEET_ID     = os.getenv("WEBENGAGE_SHEET_ID")

SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

REPORTS = [
    {"subject": "Daily WhatsApp Campaigns Summary",              "tab": "Whatsapp_Campaign"},
    {"subject": "Daily Push Campaigns Summary",                  "tab": "Push_Campaign"},
    {"subject": "Daily Journey Campaigns Summary",               "tab": "Journey_Campaign"},
    {"subject": "Daily In-app Notification Campaigns Summary",   "tab": "Inapp_Campaign"},
]


# ── Gmail auth ─────────────────────────────────────────────────────────────────

def get_gmail_service():
    creds = None
    if os.path.exists(GMAIL_TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_PATH, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow  = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_PATH, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(GMAIL_TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ── Google Sheets auth ─────────────────────────────────────────────────────────

def get_sheets_client():
    creds_path = (
        os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH")
        or os.getenv("GOOGLE_CREDENTIALS_PATH")
        or "service_account.json"
    )
    creds = ServiceCredentials.from_service_account_file(creds_path, scopes=SHEETS_SCOPES)
    return gspread.authorize(creds)


# ── Find today's email ─────────────────────────────────────────────────────────

def find_todays_email(service, subject_keyword):
    now_ist      = datetime.now(IST)
    two_days_ago = (now_ist - timedelta(days=2)).strftime("%Y/%m/%d")
    query = (
        f'from:{WEBENGAGE_SENDER} '
        f'subject:"{subject_keyword}" '
        f'after:{two_days_ago}'
    )
    result   = service.users().messages().list(userId="me", q=query, maxResults=5).execute()
    messages = result.get("messages", [])
    if not messages:
        print(f"  [Warning] No email found for: {subject_keyword}")
        return None
    return messages[0]


# ── Extract download link from email body ──────────────────────────────────────

def decode_tracking_url(link):
    """Decode base64 payload from c.webengage.com tracking URL and return toURL."""
    try:
        parsed = urllib.parse.urlparse(link)
        params = urllib.parse.parse_qs(parsed.query)
        if "p" not in params:
            return ""
        raw     = params["p"][0]
        decoded = base64.b64decode(raw + "==").decode("utf-8")
        data    = json.loads(decoded)
        return data.get("toURL", "")
    except Exception:
        return ""


def extract_download_link(service, message_id):
    msg     = service.users().messages().get(userId="me", id=message_id, format="full").execute()
    payload = msg.get("payload", {})
    parts   = payload.get("parts", [])

    html_body = ""
    if parts:
        for part in parts:
            if part.get("mimeType") == "text/html":
                html_body = base64.urlsafe_b64decode(part["body"].get("data", "")).decode("utf-8")
                break
            for sub in part.get("parts", []):
                if sub.get("mimeType") == "text/html":
                    html_body = base64.urlsafe_b64decode(sub["body"].get("data", "")).decode("utf-8")
                    break
    else:
        html_body = base64.urlsafe_b64decode(payload["body"].get("data", "")).decode("utf-8")

    # Find anchor tag with text "here" directly
    match = re.search(r'href=["\']([^"\']+)["\'][^>]*>\s*here\s*<', html_body, re.IGNORECASE)
    if match:
        link   = match.group(1)
        to_url = decode_tracking_url(link)
        print(f"  Found 'here' link -> toURL: {to_url[:80]}...")
        return link

    # Fallback: scan all c.webengage.com links, pick one pointing to S3
    links    = re.findall(r'href=["\']([^"\']+)["\']', html_body)
    we_links = [l for l in links if "c.webengage.com" in l]
    for link in we_links:
        to_url = decode_tracking_url(link)
        if "amazonaws" in to_url or "webengage-reporting" in to_url:
            print(f"  Found S3 link via fallback -> toURL: {to_url[:80]}...")
            return link

    print(f"  [Warning] No download link found in email")
    return None


# ── Download and extract CSV ───────────────────────────────────────────────────

def download_and_extract_csv(tracking_url):
    """Decode S3 URL from tracking link, download zip, extract CSV."""

    s3_url = decode_tracking_url(tracking_url)

    if not s3_url or "amazonaws" not in s3_url:
        print(f"  [Error] S3 URL not found in payload: {s3_url[:80]}")
        return []

    # Fix spaces and pipe characters in path only
    parts      = s3_url.split("?", 1)
    clean_path = parts[0].replace(" ", "%20").replace("|", "%7C")
    s3_url     = clean_path + ("?" + parts[1] if len(parts) > 1 else "")

    print(f"  Downloading from S3...")
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    response = session.get(s3_url, timeout=60)
    print(f"  S3 status: {response.status_code} | Content-Type: {response.headers.get('Content-Type', 'unknown')}")

    if response.status_code != 200:
        print(f"  [Error] S3 response: {response.content[:300]}")
        return []

    zip_bytes = io.BytesIO(response.content)
    with zipfile.ZipFile(zip_bytes, "r") as zf:
        csv_files = [f for f in zf.namelist() if f.endswith(".csv")]
        if not csv_files:
            print(f"  [Warning] No CSV found in zip")
            return []
        with zf.open(csv_files[0]) as csv_file:
            content = csv_file.read().decode("utf-8-sig")
            rows    = list(csv.reader(io.StringIO(content)))

    print(f"  Extracted {len(rows) - 1} data rows")
    return rows


# ── Write to Google Sheets ─────────────────────────────────────────────────────

def write_to_sheet(gc, tab_name, rows):
    sh = gc.open_by_key(WEBENGAGE_SHEET_ID)
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=5000, cols=50)
        print(f"  [Sheets] Created tab '{tab_name}'")
    ws.clear()
    if rows:
        ws.update(rows, value_input_option="USER_ENTERED")
    print(f"  [Sheets] Written {len(rows) - 1} rows to '{tab_name}'")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("pull_webengage_gmail.py — WebEngage Reports -> Sheets")
    print("=" * 55)

    gmail_service = get_gmail_service()
    gc            = get_sheets_client()

    for report in REPORTS:
        subject = report["subject"]
        tab     = report["tab"]
        print(f"\n[{tab}] Processing: {subject}")

        message = find_todays_email(gmail_service, subject)
        if not message:
            continue

        download_url = extract_download_link(gmail_service, message["id"])
        if not download_url:
            continue

        print(f"  Download URL found")

        rows = download_and_extract_csv(download_url)
        if not rows:
            continue

        write_to_sheet(gc, tab, rows)

    print("\n[OK] pull_webengage_gmail.py complete")


if __name__ == "__main__":
    main()