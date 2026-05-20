import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import gspread
import requests
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

META_ACCESS_TOKEN    = os.getenv("META_ACCESS_TOKEN")
META_AD_ACCOUNT_ID   = os.getenv("META_AD_ACCOUNT_ID")
GOOGLE_SHEET_ID      = os.getenv("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_PATH = os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH", "service_account.json")

META_API_VERSION  = "v19.0"
META_INSIGHTS_URL = f"https://graph.facebook.com/{META_API_VERSION}/{META_AD_ACCOUNT_ID}/insights"
META_FIELDS = (
    "campaign_name,adset_name,ad_name,objective,spend,reach,impressions,"
    "unique_outbound_clicks,unique_outbound_clicks_ctr,cpm,actions"
)

SHEET_TAB = "Meta_Ads"
HEADERS = [
    "Date", "Last_Updated", "Campaign", "Ad_Set", "Ad",
    "Objective", "Spend_INR", "Reach", "Impressions",
    "Unique_Outbound_Clicks", "Unique_Outbound_CTR",
    "Installs", "CPM", "CPI",
]

ROLLING_DAYS = 90

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

INSTALL_OBJECTIVES  = {"OUTCOME_APP_PROMOTION", "APP_INSTALLS"}
RETARGETING_MARKERS = ("ACe_", "_RT_", "_Retarget")

REQUIRED_ENV_VARS = {
    "META_ACCESS_TOKEN":  META_ACCESS_TOKEN,
    "META_AD_ACCOUNT_ID": META_AD_ACCOUNT_ID,
    "GOOGLE_SHEET_ID":    GOOGLE_SHEET_ID,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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


def _action_value(items: list, action_type: str) -> str:
    """Extract the value for a specific action_type from a Meta actions array."""
    for item in items or []:
        if item.get("action_type") == action_type:
            return item.get("value", "")
    return ""


def _fmt_float(value, decimals: int = 2) -> str:
    try:
        return f"{float(value):.{decimals}f}"
    except (TypeError, ValueError):
        return ""


# ---------------------------------------------------------------------------
# Meta API
# ---------------------------------------------------------------------------

def _get_with_retry(url: str, params: dict, max_retries: int = 4) -> requests.Response:
    """GET with exponential backoff on Meta rate-limit responses."""
    rate_limit_codes = {4, 17, 613, 80004}
    for attempt in range(max_retries):
        response = requests.get(url, params=params, timeout=60)
        if response.status_code == 429:
            wait = 10 * (2 ** attempt)
            print(f"  Rate limited (HTTP 429), waiting {wait}s...")
            time.sleep(wait)
            continue
        if response.status_code in (400, 500):
            try:
                code = response.json().get("error", {}).get("code", 0)
            except Exception:
                code = 0
            if code in rate_limit_codes:
                wait = 10 * (2 ** attempt)
                print(f"  Rate limited (error code {code}), waiting {wait}s...")
                time.sleep(wait)
                continue
        return response
    return response


def fetch_meta_insights(date_str: str) -> list[dict]:
    """
    Fetch ad-level insights for a single day with full pagination.
    Date is stamped manually by the caller — never read from API rows.
    """
    params = {
        "access_token": META_ACCESS_TOKEN,
        "level":        "ad",
        "fields":       META_FIELDS,
        "time_range":   json.dumps({"since": date_str, "until": date_str}),
        "limit":        500,
    }

    all_data: list[dict] = []
    current_url    = META_INSIGHTS_URL
    current_params = params

    while True:
        response = _get_with_retry(current_url, current_params)
        if not response.ok:
            try:
                err = response.json().get("error", {}).get("message", response.text)
            except Exception:
                err = response.text
            raise RuntimeError(f"Meta API error {response.status_code}: {err}")

        body = response.json()
        all_data.extend(body.get("data", []))

        # paging.next already has all params encoded — pass empty params
        next_url = body.get("paging", {}).get("next")
        if not next_url:
            break
        current_url    = next_url
        current_params = {}

    return all_data


def normalize_row(raw: dict) -> dict:
    """Map a raw Meta API row to our schema."""
    campaign    = raw.get("campaign_name", "")
    adset       = raw.get("adset_name", "")
    ad          = raw.get("ad_name", "")
    objective   = raw.get("objective", "")
    spend       = raw.get("spend", "0")
    reach       = raw.get("reach", "")
    impressions = raw.get("impressions", "")
    cpm         = _fmt_float(raw.get("cpm"))

    # Unique outbound clicks — unique_outbound_clicks array, action_type = outbound_click
    unique_outbound_clicks = _action_value(
        raw.get("unique_outbound_clicks", []), "outbound_click"
    )

    # Unique outbound CTR — unique_outbound_clicks_ctr array, action_type = outbound_click
    unique_outbound_ctr = _fmt_float(
        _action_value(raw.get("unique_outbound_clicks_ctr", []), "outbound_click")
    )

    # Installs — actions array, action_type = mobile_app_install
    installs = _action_value(raw.get("actions", []), "mobile_app_install")

    # CPI — only for install objectives on non-retargeting campaigns
    is_install_obj = any(obj in objective.upper() for obj in INSTALL_OBJECTIVES)
    is_retargeting = any(marker in campaign for marker in RETARGETING_MARKERS)

    if is_install_obj and not is_retargeting:
        try:
            install_count = float(installs) if installs else 0
            cpi = f"{float(spend) / install_count:.2f}" if install_count > 0 else ""
        except (ValueError, ZeroDivisionError):
            cpi = ""
    else:
        cpi = "N/A"

    return {
        "campaign":               campaign,
        "adset":                  adset,
        "ad":                     ad,
        "objective":              objective,
        "spend_inr":              _fmt_float(spend),
        "reach":                  reach,
        "impressions":            impressions,
        "unique_outbound_clicks": unique_outbound_clicks,
        "unique_outbound_ctr":    unique_outbound_ctr,
        "installs":               installs,
        "cpm":                    cpm,
        "cpi":                    cpi,
    }


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------

def connect_sheet() -> gspread.Worksheet:
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
    try:
        worksheet = spreadsheet.worksheet(SHEET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=SHEET_TAB, rows=5000, cols=len(HEADERS)
        )
    return worksheet


def ensure_headers(worksheet: gspread.Worksheet) -> None:
    if worksheet.row_values(1) != HEADERS:
        worksheet.update("A1", [HEADERS])


def purge_old_rows(worksheet: gspread.Worksheet, cutoff_date: str) -> int:
    all_rows = worksheet.get_all_values()
    if len(all_rows) <= 1:
        return 0
    to_delete = [
        i + 2
        for i, row in enumerate(all_rows[1:])
        if row and row[0] and row[0] < cutoff_date
    ]
    for row_num in reversed(to_delete):
        worksheet.delete_rows(row_num)
    return len(to_delete)


def unique_key(row_values: list) -> tuple:
    """Date + Campaign + Ad_Set + Ad."""
    idx = {h: i for i, h in enumerate(HEADERS)}
    return (
        row_values[idx["Date"]],
        row_values[idx["Campaign"]],
        row_values[idx["Ad_Set"]],
        row_values[idx["Ad"]],
    )


def write_rows(worksheet: gspread.Worksheet, new_rows: list[list]) -> tuple[int, int]:
    """Batch upsert: overwrite matched rows, append new ones."""
    all_data = worksheet.get_all_values()
    existing = all_data[1:] if len(all_data) > 1 else []

    lookup: dict[tuple, int] = {}
    for i, row in enumerate(existing):
        padded = row + [""] * (len(HEADERS) - len(row))
        lookup[unique_key(padded)] = i + 2

    batch_updates = []
    to_insert     = []

    for row_values in new_rows:
        key = unique_key(row_values)
        if key in lookup:
            batch_updates.append({
                "range":  f"A{lookup[key]}",
                "values": [row_values],
            })
        else:
            to_insert.append(row_values)

    if batch_updates:
        worksheet.batch_update(batch_updates, value_input_option="USER_ENTERED")
    if to_insert:
        worksheet.append_rows(to_insert, value_input_option="USER_ENTERED")

    return len(batch_updates), len(to_insert)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    missing = [k for k, v in REQUIRED_ENV_VARS.items() if not v]
    if missing:
        for var in missing:
            print(f"ERROR: {var} not set in environment")
        sys.exit(1)

    now_str     = get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
    today       = get_today_ist()
    yesterday   = get_yesterday_ist()
    month_start = get_month_start_ist()
    cutoff_date = (
        datetime.strptime(yesterday, "%Y-%m-%d") - timedelta(days=ROLLING_DAYS - 1)
    ).strftime("%Y-%m-%d")

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
        print(f"Empty sheet — backfilling Meta Ads data from {month_start} to {today} ({len(date_set)} days)...")
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
        print(f"Fetching Meta Ads data — last date in sheet: {last_date}, dates to process: {len(date_set)}...")

    date_range: list[str] = sorted(date_set)
    label = date_range[0] if len(date_range) == 1 else f"{date_range[0]} to {date_range[-1]}"
    print(f"Processing {len(date_range)} day(s): {label}")

    all_sheet_rows: list[list] = []
    total_raw = 0

    for i, day in enumerate(date_range):
        try:
            raw_rows = fetch_meta_insights(day)
        except RuntimeError as e:
            print(f"ERROR: Meta API request failed for {day} — {e}")
            sys.exit(1)
        except Exception as e:
            print(f"ERROR: Meta API request failed for {day} — {e}")
            sys.exit(1)

        total_raw += len(raw_rows)

        for raw in raw_rows:
            row = normalize_row(raw)
            all_sheet_rows.append([
                day,
                now_str,
                row["campaign"],
                row["adset"],
                row["ad"],
                row["objective"],
                row["spend_inr"],
                row["reach"],
                row["impressions"],
                row["unique_outbound_clicks"],
                row["unique_outbound_ctr"],
                row["installs"],
                row["cpm"],
                row["cpi"],
            ])

        # 1 second between day iterations to respect Meta rate limits
        if i < len(date_range) - 1:
            time.sleep(1)

    print(f"  Raw rows received: {total_raw}")
    print(f"  Sheet rows to write: {len(all_sheet_rows)}")

    try:
        deleted = purge_old_rows(worksheet, cutoff_date)
        if deleted:
            print(f"  Purged {deleted} rows older than {ROLLING_DAYS} days")
    except Exception as e:
        print(f"WARNING: Failed to purge old rows — {e}")

    if not all_sheet_rows:
        print("No data to write.")
        print(f"SUCCESS: Meta Ads pull complete ({label})")
        return

    try:
        updated, inserted = write_rows(worksheet, all_sheet_rows)
    except Exception as e:
        print(f"ERROR: Failed to write to Google Sheets — {e}")
        sys.exit(1)

    print(f"  Rows updated: {updated} | Rows inserted: {inserted}")
    print(f"SUCCESS: Meta Ads pull complete ({label})")


if __name__ == "__main__":
    main()
