import os
import sys
import requests
from datetime import datetime, timedelta, timezone

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

APPTROVE_API_KEY = os.getenv("APPTROVE_MMP_API_KEY")
APPTROVE_APP_ID = os.getenv("APPTROVE_APP_ID")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_PATH = os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH", "service_account.json")

SHEET_TAB = "Apptrove_MMP"
HEADERS = [
    "Date", "Last_Updated", "partner", "channel", "campaign",
    "ad_group", "ad", "app_opened", "first_homePage_viewed",
    "view_content", "purchase", "first_purchase_success",
]

APPTROVE_API_URL = f"https://api.apptrove.com/api/v1/app/{APPTROVE_APP_ID}/report/event"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

# Event IDs for Unique Events Accepted — keyed by our column name
EVENT_IDS = {
    "app_opened":              "bSTMxdlv3O",
    "first_homePage_viewed":   "bs5Ta9Lwed",
    "view_content":            "wutNNRc73E",
    "purchase":                "owi4J2JMOB",
    "first_purchase_success":  "dLoYBlZhbw",
}

# Reverse map: event ID -> column name (for parsing the API response)
EVENT_ID_TO_NAME = {v: k for k, v in EVENT_IDS.items()}

GOOGLE_PARTNER_NAME = "Google Ads (Adwords)"
# Apptrove may surface Meta under several names; add variants as needed
META_PARTNER_NAMES = {"Facebook Ads", "Meta Ads", "Facebook"}

INVALID_CHANNEL = "-"
ROLLING_DAYS = 90


def get_ist_now() -> datetime:
    return datetime.now(timezone(timedelta(hours=5, minutes=30)))


def get_yesterday_ist() -> str:
    return (get_ist_now() - timedelta(days=1)).strftime("%Y-%m-%d")


def get_today_ist() -> str:
    return get_ist_now().strftime("%Y-%m-%d")


def get_month_start_ist() -> str:
    """Return the 1st of the current IST month as YYYY-MM-DD."""
    return get_ist_now().replace(day=1).strftime("%Y-%m-%d")


def get_existing_dates(worksheet: gspread.Worksheet) -> set[str]:
    """Return the set of date strings found in column A (excluding header)."""
    all_values = worksheet.get_all_values()
    if len(all_values) <= 1:
        return set()
    return {row[0] for row in all_values[1:] if row and row[0]}


def fetch_apptrove_data(start_date: str, end_date: str) -> list[dict]:
    """
    GET Apptrove event report for a date range.

    Response shape: {"success": true, "data": {"records": [...], ...}}
    Each record is one event for one dimension combination, e.g.:
      {"Partner": "Google Ads (Adwords)", "Channel": "Search",
       "Campaign": "...", "Adset": "...", "Ad": null,
       "Event": "app_opened", "Event ID": "bSTMxdlv3O",
       "Uevents": 90, "date": "2026-05-10"}

    date is included via groupBy[] so records carry their own date.
    Uevents = Unique Events Accepted (what we always want).
    """
    headers = {
        "api-key": APPTROVE_API_KEY,
        "Accept": "application/json",
    }
    params = [
        ("start_date", start_date),
        ("end_date", end_date),
        ("event_counting", "unique"),
        ("groupBy[]", "partner"),
        ("groupBy[]", "channel"),
        ("groupBy[]", "campaign"),
        ("groupBy[]", "adset"),
        ("groupBy[]", "ad"),
    ]
    for eid in EVENT_IDS.values():
        params.append(("eid[]", eid))

    response = requests.get(APPTROVE_API_URL, headers=headers, params=params, timeout=60)
    if not response.ok:
        print(f"  [Apptrove] Status: {response.status_code}")
        print(f"  [Apptrove] Response body: {response.text[:500]}")
    response.raise_for_status()

    body = response.json()
    records = body.get("data", {}).get("records")
    if not isinstance(records, list):
        raise ValueError(f"Expected data.records to be a list, got: {type(records)}")
    return records


# EVENT_NAME_TO_COL: matches the "Event" field value in each record
EVENT_NAME_TO_COL = {
    "app_opened":            "app_opened",
    "first_homePage_viewed": "first_homePage_viewed",
    "view_content":          "view_content",
    "purchase":              "purchase",
    "first_purchase_success":"first_purchase_success",
}


def pivot_records(records: list[dict], fallback_date: str) -> list[dict]:
    """
    The API returns one record per (date, partner, channel, campaign, adset, ad, event).
    Pivot to one row per (date, partner, channel, campaign, adset, ad) with all
    event counts as columns.  Apply partner-specific field rules and skip
    Google rows where Channel == "-".
    """
    # Use an ordered dict keyed by the group tuple to accumulate event counts
    groups: dict[tuple, dict] = {}

    for rec in records:
        partner  = (rec.get("Partner")  or "").strip()
        channel  = (rec.get("Channel")  or "").strip()
        campaign = (rec.get("Campaign") or "").strip()
        adset    = (rec.get("Adset")    or "").strip()
        ad_raw   = (rec.get("Ad")       or "").strip()
        row_date = fallback_date

        # Skip invalid Google channel entries
        if partner == GOOGLE_PARTNER_NAME and channel == INVALID_CHANNEL:
            continue

        is_google = partner == GOOGLE_PARTNER_NAME
        is_meta   = partner in META_PARTNER_NAMES

        if is_google:
            ad_group = adset
            ad       = ""
        elif is_meta:
            ad_group = ""
            channel  = ""
            ad       = ad_raw
        else:
            ad_group = adset
            ad       = ad_raw

        group_key = (row_date, partner, channel, campaign, ad_group, ad)

        if group_key not in groups:
            groups[group_key] = {
                "date": row_date, "partner": partner, "channel": channel,
                "campaign": campaign, "ad_group": ad_group, "ad": ad,
                **{col: "" for col in EVENT_IDS},
            }

        # Map the event name in this record to our column and store Uevents
        event_name = (rec.get("Event") or "").strip()
        event_id   = (rec.get("Event ID") or "").strip()
        col = EVENT_NAME_TO_COL.get(event_name) or EVENT_ID_TO_NAME.get(event_id)
        if col:
            groups[group_key][col] = _safe_int(rec.get("Uevents"))

    return list(groups.values())


def _safe_int(value) -> str:
    """Return integer string or empty string for missing/None values."""
    if value is None or value == "":
        return ""
    try:
        return str(int(value))
    except (ValueError, TypeError):
        return str(value)


def unique_key(row_values: list, date_col: int = 0) -> tuple:
    """
    Unique identifier: Date + partner + channel + campaign + ad_group.
    Expects row_values aligned to HEADERS.
    """
    idx = {h: i for i, h in enumerate(HEADERS)}
    return (
        row_values[idx["Date"]],
        row_values[idx["partner"]],
        row_values[idx["channel"]],
        row_values[idx["campaign"]],
        row_values[idx["ad_group"]],
    )


def connect_sheet() -> gspread.Worksheet:
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
    try:
        worksheet = spreadsheet.worksheet(SHEET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=SHEET_TAB, rows=1000, cols=len(HEADERS))
    return worksheet


def ensure_headers(worksheet: gspread.Worksheet) -> None:
    existing = worksheet.row_values(1)
    if existing != HEADERS:
        worksheet.update("A1", [HEADERS])


def purge_old_rows(worksheet: gspread.Worksheet, cutoff_date: str) -> int:
    """Delete rows whose Date is older than cutoff_date. Returns count deleted."""
    all_rows = worksheet.get_all_values()
    if len(all_rows) <= 1:
        return 0

    rows_to_delete = []
    for i, row in enumerate(all_rows[1:], start=2):  # 1-indexed, skip header
        if row and row[0] and row[0] < cutoff_date:
            rows_to_delete.append(i)

    # Delete bottom-up so row indices remain valid
    for row_num in reversed(rows_to_delete):
        worksheet.delete_rows(row_num)

    return len(rows_to_delete)


def write_rows(worksheet: gspread.Worksheet, new_rows: list[list], date_str: str) -> tuple[int, int]:
    """
    Overwrite existing rows matched by unique key; append new ones.
    Uses batch operations to stay within Sheets API rate limits.
    Returns (updated_count, inserted_count).
    """
    all_data = worksheet.get_all_values()
    existing_rows = all_data[1:] if len(all_data) > 1 else []

    lookup: dict[tuple, int] = {}
    for i, row in enumerate(existing_rows):
        padded = row + [""] * (len(HEADERS) - len(row))
        key = unique_key(padded)
        lookup[key] = i + 2  # 1-indexed sheet row, header is row 1

    batch_updates = []  # list of {"range": "A{n}", "values": [[...]]}
    rows_to_insert = []

    for row_values in new_rows:
        key = unique_key(row_values)
        if key in lookup:
            sheet_row = lookup[key]
            batch_updates.append({
                "range": f"A{sheet_row}",
                "values": [row_values],
            })
        else:
            rows_to_insert.append(row_values)

    # Single batch_update call for all overwrites
    if batch_updates:
        worksheet.batch_update(batch_updates, value_input_option="USER_ENTERED")

    # Single append_rows call for all inserts
    if rows_to_insert:
        worksheet.append_rows(rows_to_insert, value_input_option="USER_ENTERED")

    return len(batch_updates), len(rows_to_insert)


def main():
    if not APPTROVE_API_KEY:
        print("ERROR: APPTROVE_MMP_API_KEY not set in environment")
        sys.exit(1)
    if not APPTROVE_APP_ID:
        print("ERROR: APPTROVE_APP_ID not set in environment")
        sys.exit(1)
    if not GOOGLE_SHEET_ID:
        print("ERROR: GOOGLE_SHEET_ID not set in environment")
        sys.exit(1)

    now_str = get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
    today = get_today_ist()
    yesterday = get_yesterday_ist()
    month_start = get_month_start_ist()
    cutoff_date = (datetime.strptime(yesterday, "%Y-%m-%d") - timedelta(days=ROLLING_DAYS - 1)).strftime("%Y-%m-%d")

    try:
        worksheet = connect_sheet()
        ensure_headers(worksheet)
    except Exception as e:
        print(f"ERROR: Failed to connect to Google Sheets — {e}")
        sys.exit(1)

    existing_dates = get_existing_dates(worksheet)

    if not existing_dates:
        # Empty sheet: backfill from 1st of current month through today
        date_set: set[str] = set()
        d = datetime.strptime(month_start, "%Y-%m-%d")
        end_d = datetime.strptime(today, "%Y-%m-%d")
        while d <= end_d:
            date_set.add(d.strftime("%Y-%m-%d"))
            d += timedelta(days=1)
        print(f"Empty sheet — backfilling Apptrove MMP data from {month_start} to {today} ({len(date_set)} days)...")
    else:
        last_date = max(existing_dates)
        # Gap fill: last_date+1 through yesterday, then add yesterday and today
        date_set = set()
        d = datetime.strptime(last_date, "%Y-%m-%d") + timedelta(days=1)
        end_d = datetime.strptime(yesterday, "%Y-%m-%d")
        while d <= end_d:
            date_set.add(d.strftime("%Y-%m-%d"))
            d += timedelta(days=1)
        date_set.add(yesterday)
        date_set.add(today)
        print(f"Fetching Apptrove MMP data — last date in sheet: {last_date}, dates to process: {len(date_set)}...")

    date_range = sorted(date_set)

    sheet_rows = []
    total_raw = 0
    for day in date_range:
        try:
            raw_rows = fetch_apptrove_data(day, day)
        except requests.HTTPError as e:
            print(f"ERROR: Apptrove API request failed for {day} — {e}")
            sys.exit(1)
        except Exception as e:
            print(f"ERROR: Failed to fetch Apptrove data for {day} — {e}")
            sys.exit(1)

        total_raw += len(raw_rows)
        for row in pivot_records(raw_rows, fallback_date=day):
            sheet_rows.append([
                row["date"],
                now_str,
                row["partner"],
                row["channel"],
                row["campaign"],
                row["ad_group"],
                row["ad"],
                row["app_opened"],
                row["first_homePage_viewed"],
                row["view_content"],
                row["purchase"],
                row["first_purchase_success"],
            ])

    print(f"  Raw records received: {total_raw}")
    print(f"  Rows after pivot & filter: {len(sheet_rows)}")

    try:
        deleted = purge_old_rows(worksheet, cutoff_date)
        if deleted:
            print(f"  Purged {deleted} rows older than {ROLLING_DAYS} days")
    except Exception as e:
        print(f"WARNING: Failed to purge old rows — {e}")

    if not sheet_rows:
        print("No data to write.")
        print("SUCCESS: Apptrove MMP pull complete (no rows written)")
        return

    try:
        updated, inserted = write_rows(worksheet, sheet_rows, today)
    except Exception as e:
        print(f"ERROR: Failed to write to Google Sheets — {e}")
        sys.exit(1)

    print(f"  Rows updated: {updated} | Rows inserted: {inserted}")
    label = date_range[0] if len(date_range) == 1 else f"{date_range[0]} to {date_range[-1]}"
    print(f"SUCCESS: Apptrove MMP pull complete ({label})")


if __name__ == "__main__":
    main()