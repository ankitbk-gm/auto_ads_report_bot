import os
import requests
from datetime import datetime, timedelta, timezone, date
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

IST = ZoneInfo("Asia/Kolkata")
UTC = timezone.utc

WATI_ACCOUNT_ID = os.getenv("WATI_ACCOUNT_ID", "455444")
WATI_BASE        = f"https://live-mt-server.wati.io/{WATI_ACCOUNT_ID}/api/v1"
WATI_TOKEN = os.getenv("WATI_API_TOKEN")
HEADERS    = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json"}

CHANNELS = [
    "685e3e658e287bf95a1ca44a",  # +919910229390
    "685e3de3c109042bc410aac8",  # +919871200496
    "6996b82ae2494d340cf7279b",  # +919810154799
    "699eb4eaa826c8cacecb1188",  # +919870409158
]

SHEET_TAB = "Wati"
COLUMNS   = ["Date", "Last_Updated", "Campaign_Name", "Template_Name",
             "Sent", "Failed", "Delivered", "Read"]


# ── Sheets ─────────────────────────────────────────────────────────────────────

def get_sheet():
    creds = Credentials.from_service_account_file(
        os.getenv("GOOGLE_CREDENTIALS_PATH") or os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH"),
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds).open_by_key(
        os.getenv("GOOGLE_SHEET_ID")).worksheet(SHEET_TAB)


def is_initial_run(ws):
    return len(ws.get_all_values()) <= 1


def get_existing_index(ws):
    rows = ws.get_all_values()
    if len(rows) <= 1:
        return {}
    header = rows[0]
    index  = {}
    for i, row in enumerate(rows[1:], start=2):
        r   = dict(zip(header, row))
        key = (r.get("Date",""), r.get("Campaign_Name",""), r.get("Template_Name",""))
        index[key] = i
    return index


def get_latest_date(ws):
    """Returns the latest date found in the sheet as a date object, or None."""
    rows = ws.get_all_values()
    if len(rows) <= 1:
        return None
    latest = None
    for row in rows[1:]:
        if row and row[0]:
            try:
                d = datetime.strptime(row[0], "%d-%m-%Y").date()
                if latest is None or d > latest:
                    latest = d
            except ValueError:
                continue
    return latest


# ── API: 4 channels ────────────────────────────────────────────────────────────

def fetch_for_channel(channel_id, date_str, partial=False):
    url     = f"{WATI_BASE}/broadcast?ChannelId={channel_id}"
    date_to = (datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S.999Z")
               if partial else f"{date_str}T23:59:59.999Z")
    items   = []
    page    = 0
    while True:
        body = {
            "dateFrom":   f"{date_str}T00:00:00.000Z",
            "dateTo":     date_to,
            "pageSize":   50,
            "pageNumber": page,
            "isUpdate":   False,
        }
        try:
            r = requests.post(url, headers=HEADERS, json=body, timeout=30)
            if r.status_code != 200:
                print(f"  [{channel_id[:8]}] API {r.status_code}: {r.text[:100]}")
                break
            batch = (r.json().get("result") or {}).get("items") or []
            items.extend(batch)
            if len(batch) < 50:
                break
            page += 1
        except Exception as e:
            print(f"  [{channel_id[:8]}] Error: {e}")
            break
    return items


# ── API: default channel overview (voice_ai) ───────────────────────────────────

def fetch_default_overview(date_str, partial=False):
    date_to = (datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S.999Z")
               if partial else f"{date_str}T23:59:59.999Z")
    url  = f"{WATI_BASE}/broadcast/getBroadcastsOverview"
    body = {
        "dateFrom": f"{date_str}T00:00:00.000Z",
        "dateTo":   date_to,
    }
    try:
        r = requests.post(url, headers=HEADERS, json=body, timeout=30)
        if r.status_code != 200:
            print(f"  [default] getBroadcastsOverview {r.status_code}: {r.text[:100]}")
            return None
        result = (r.json().get("result") or {})
        return {
            "Sent":      result.get("totalSent",      0) or 0,
            "Failed":    result.get("totalFailed",    0) or 0,
            "Delivered": result.get("totalDelivered", 0) or 0,
            "Read":      result.get("totalOpen",      0) or 0,
        }
    except Exception as e:
        print(f"  [default] Error: {e}")
        return None


# ── Process one date ───────────────────────────────────────────────────────────

def process_date(target_date, partial=False):
    date_str    = target_date.strftime("%Y-%m-%d")
    display_str = target_date.strftime("%d-%m-%Y")
    now_ist_str = datetime.now(IST).strftime("%d-%m-%Y %H:%M:%S")

    agg = {}
    for cid in CHANNELS:
        items = fetch_for_channel(cid, date_str, partial)
        print(f"  [{display_str}] [{cid[:8]}] -> {len(items)} campaigns")
        for item in items:
            try:
                sat = item.get("scheduledAt", "")
                if datetime.fromisoformat(sat.replace("Z", "+00:00")).strftime("%Y-%m-%d") != date_str:
                    continue
            except:
                continue

            k = (item.get("broadcastName", ""), item.get("templateName", ""))
            if k not in agg:
                agg[k] = {
                    "Campaign_Name": k[0],
                    "Template_Name": k[1],
                    "Sent": 0, "Failed": 0, "Delivered": 0, "Read": 0,
                }
            agg[k]["Sent"]      += item.get("recipients",     0) or 0
            agg[k]["Failed"]    += item.get("fail",           0) or 0
            agg[k]["Delivered"] += item.get("deliveredCount", 0) or 0
            agg[k]["Read"]      += item.get("readCount",      0) or 0

    print(f"  [{display_str}] {len(agg)} unique campaigns from 4 channels")

    rows = [{
        "Date":          display_str,
        "Last_Updated":  now_ist_str,
        "Campaign_Name": v["Campaign_Name"],
        "Template_Name": v["Template_Name"],
        "Sent":          v["Sent"],
        "Failed":        v["Failed"],
        "Delivered":     v["Delivered"],
        "Read":          v["Read"],
    } for v in agg.values()]

    # Default channel: voice_ai row
    overview = fetch_default_overview(date_str, partial)
    if overview and any(overview.values()):
        rows.append({
            "Date":          display_str,
            "Last_Updated":  now_ist_str,
            "Campaign_Name": "voice_ai",
            "Template_Name": "voice_ai",
            "Sent":          overview["Sent"],
            "Failed":        overview["Failed"],
            "Delivered":     overview["Delivered"],
            "Read":          overview["Read"],
        })
        print(f"  [{display_str}] voice_ai row added")

    return rows


# ── Sheets write ───────────────────────────────────────────────────────────────

def upsert_rows(ws, rows, idx):
    updates, appends = [], []
    for row in rows:
        key    = (row["Date"], row["Campaign_Name"], row["Template_Name"])
        values = [row[c] for c in COLUMNS]
        if key in idx:
            updates.append((idx[key], values))
        else:
            appends.append(values)

    if updates:
        col = chr(ord("A") + len(COLUMNS) - 1)
        ws.spreadsheet.values_batch_update({
            "valueInputOption": "RAW",
            "data": [
                {"range": f"{SHEET_TAB}!A{n}:{col}{n}", "values": [v]}
                for n, v in updates
            ]
        })
    if appends:
        ws.append_rows(appends, value_input_option="RAW")

    print(f"  Overwritten: {len(updates)} | Appended: {len(appends)}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=== pull_wati.py ===")

    ws      = get_sheet()
    now_ist = datetime.now(IST)
    today     = now_ist.replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday = today - timedelta(days=1)

    # ── Initial run ────────────────────────────────────────────────────────────
    if is_initial_run(ws):
        month_start = now_ist.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        print(f"Initial run: {month_start.strftime('%d-%m-%Y')} to {today.strftime('%d-%m-%Y')}")
        rows, cur = [], month_start
        while cur < today:
            rows.extend(process_date(cur, partial=False))
            cur += timedelta(days=1)
        rows.extend(process_date(today, partial=True))
        ws.clear()
        ws.append_row(COLUMNS)
        if rows:
            ws.append_rows(
                [[r[c] for c in COLUMNS] for r in rows],
                value_input_option="RAW"
            )
        print(f"SUCCESS -- {len(rows)} rows written (initial run)")
        return

    # ── Regular run ────────────────────────────────────────────────────────────
    idx         = get_existing_index(ws)
    latest_date = get_latest_date(ws)
    rows        = []

    if latest_date and latest_date < yesterday.date():
        # Gap detected — backfill from latest + 1 to yesterday
        backfill_start = datetime.combine(
            latest_date + timedelta(days=1),
            datetime.min.time()
        ).replace(tzinfo=IST)
        print(f"Gap detected. Backfilling: {backfill_start.strftime('%d-%m-%Y')} → {yesterday.strftime('%d-%m-%Y')}")
        cur = backfill_start
        while cur.date() <= yesterday.date():
            rows.extend(process_date(cur, partial=False))
            cur += timedelta(days=1)
    else:
        # No gap — just pull yesterday
        print(f"Yesterday: {yesterday.strftime('%d-%m-%Y')}")
        rows.extend(process_date(yesterday, partial=False))

    # 6pm run: also refresh today partial
    if now_ist.hour >= 12:
        print(f"Today partial: {today.strftime('%d-%m-%Y')}")
        rows.extend(process_date(today, partial=True))

    upsert_rows(ws, rows, idx)
    print(f"SUCCESS -- {len(rows)} rows processed (regular run)")


if __name__ == "__main__":
    main()