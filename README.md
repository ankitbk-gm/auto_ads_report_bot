# auto_ads_report_bot
Automated ads reporting — Google, Meta, Trackier (MMP), Webengage. Purpose: To analyze full funnel marketing for an App
# Auto Ads Report Bot

Automated marketing data pipeline for Agrim Wholesale — pulls data from Google Ads, Meta Ads, Wati (WhatsApp), WebEngage and Apptrove MMP into Google Sheets and OneDrive Excel reports.

---

## Architecture

```
Google Ads (Script)  ──┐
Meta Ads API         ──┤
Wati API             ──┼──► Google Sheets (Raw Data) ──► Report Scripts ──► Report Sheet + OneDrive Excel
Apptrove MMP API     ──┤
WebEngage (Gmail)    ──┘
```

---

## Scripts

### Data Pull Scripts (`src/`)

| Script | Source | Destination | Description |
|---|---|---|---|
| `pull_meta.py` | Meta Ads API | `Meta_Ads` tab | Daily Meta campaign spend, impressions, reach |
| `pull_apptrove_mmp.py` | Apptrove API | `Apptrove_MMP` tab | MMP attribution events (installs, conversions) |
| `pull_wati.py` | Wati API | `Wati` tab | WhatsApp campaign delivery stats with gap detection |
| `pull_webengage_gmail.py` | Gmail (WebEngage emails) | `Whatsapp_Campaign`, `Push_Campaign`, `Journey_Campaign`, `Inapp_Campaign` tabs | Fetches WebEngage daily report CSVs from Gmail |

### Report Scripts (`src/`)

| Script | Output Tab | Description |
|---|---|---|
| `pull_team_cost_report.py` | `Google_Campaigns`, `Meta_Campaigns`, `Wati_Cost`, `WebEngage_Cost`, `Team_Wise_Cost` | Daily cost breakdown by team (MTD / MTD-1 / Change) |
| `pull_onboarding_report.py` | `Onboarding_report` | KYC and First Transaction funnel by channel (Google / Meta / WhatsApp) |
| `pull_retention_report.py` | `retention_brand_report` | Retention campaign performance (Brand vs Subcat) for Google and Meta |
| `pull_kam_brand_report.py` | `KAM_Brand_Report` | Brand-level performance with reach, impressions, spend, app traffic |
| `pull_inapp_subcat_report.py` | `Inapp_Subcat_Report` | In-app subcategory promotion report with persistent monthly storage |
| `pull_webengage_analysis.py` | `Webengage_Analysis` | Channel performance, team engagement, journey funnel, WA failure analysis, push engagement |

---

## Run Order

```bash
py src/pull_meta.py
py src/pull_apptrove_mmp.py
py src/pull_wati.py
py src/pull_webengage_gmail.py
py src/pull_team_cost_report.py
py src/pull_onboarding_report.py
py src/pull_retention_report.py
py src/pull_kam_brand_report.py
py src/pull_inapp_subcat_report.py
py src/pull_webengage_analysis.py
```

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/ankitbk-gm/auto_ads_report_bot.git
cd auto_ads_report_bot
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure environment variables

Copy `.env.example` to `.env` and fill in all values:

```bash
cp .env.example .env
```

### 4. Add credential files

Place the following files in the project root (never commit these):
- `service_account.json` — Google Service Account for Sheets access
- `gmail_credentials.json` — Gmail OAuth credentials
- `gmail_token.json` — Gmail OAuth token (auto-generated on first run)

### 5. Run

```bash
py src/pull_meta.py
```

---

## Environment Variables

See `.env.example` for the full list. Key variables:

| Variable | Description |
|---|---|
| `GOOGLE_SHEET_ID` | Raw data Google Sheet ID |
| `REPORT_SHEET_ID` | Report output Google Sheet ID |
| `WEBENGAGE_SHEET_ID` | WebEngage reports Google Sheet ID |
| `GOOGLE_UNIQUE_USERS_SHEET_ID` | Google Ads unique users scheduled report sheet |
| `META_ACCESS_TOKEN` | Meta Ads API access token |
| `META_AD_ACCOUNT_ID` | Meta Ad account ID (format: act_XXXXXXXXXX) |
| `WATI_API_TOKEN` | Wati API bearer token |
| `WATI_ACCOUNT_ID` | Wati account ID (numeric) |
| `GMAIL_CREDENTIALS_PATH` | Path to Gmail OAuth credentials file |
| `GMAIL_TOKEN_PATH` | Path to Gmail OAuth token file |

---

## Google Sheets Structure

### Raw Data Sheet (`GOOGLE_SHEET_ID`)
- `Google_Ads` — daily campaign + ad group level data
- `Meta_Ads` — daily campaign + ad level data
- `Wati` — WhatsApp campaign delivery data
- `Apptrove_MMP` — MMP attribution events by partner + campaign

### WebEngage Sheet (`WEBENGAGE_SHEET_ID`)
- `Whatsapp_Campaign` — last 30 days WhatsApp campaign stats
- `Push_Campaign` — last 30 days Push campaign stats
- `Journey_Campaign` — last 30 days Journey stats
- `Inapp_Campaign` — last 30 days In-app campaign stats

### Report Sheet (`REPORT_SHEET_ID`)
- `Google_Campaigns`, `Meta_Campaigns`, `Wati_Cost`, `WebEngage_Cost`
- `Team_Wise_Cost` — MTD / MTD-1 / Change by team
- `Onboarding_report` — KYC + First Txn funnel
- `retention_brand_report` — Brand vs Subcat retention
- `KAM_Brand_Report` — Brand performance
- `Inapp_Subcat_Report` — In-app subcategory report
- `Inapp_Subcat_Store` — Persistent monthly storage for in-app data
- `Webengage_Analysis` — Full WebEngage analysis

---

## Team Mapping Logic

### Wati
| Keyword | Team |
|---|---|
| `voice_ai` | Voice AI |
| `ss_*` | Superstar |
| `customer_support`, `bot_` | Customer Support |
| `marketing_onboarding` | Marketing Onboarding |
| `marketing_retention`, `marketing_*` | Marketing Retention |

### Meta
| Keyword | Team |
|---|---|
| `superstar`, `ss_` | Superstar |
| `seller` | Supply |
| `ar_*install`, `onboarding`, `aci`, `otp_entered`, `kyced` | Marketing Onboarding |
| `retention`, `ace` | Marketing Retention |

### Google
| Keyword | Team |
|---|---|
| `aci`, `otp_entered`, `kyced`, `install`, `onboarding` | Marketing Onboarding |
| `retention`, `ace` | Marketing Retention |

### WebEngage
| Type | Condition | Team |
|---|---|---|
| Relay | — | Marketing Retention |
| Journey | add to cart | Marketing Retention |
| Journey | other | Marketing Onboarding |
| One-time | onboarding in name | Marketing Onboarding |
| One-time | other | Marketing Retention |

---

## Naming Conventions

For reports to work correctly, campaigns must follow these naming conventions:

### Google Ads
- Brand campaigns: `Brand_Promotion_{BrandName}` (ad group level)
- Category campaigns: `Category_Promotion_{CategoryName}`
- KYC campaigns: include `otp_entered` or `aci`
- TXN campaigns: include `kyced_but_not_transacted`
- Retention: include `ace` or `ar_purchasers`

### Meta Ads
- Brand ads: `{BrandName}_{Date}` (e.g. `HPM_7Apr26`)
- Subcat ads: include `catalogue`, `category`, or `gibberellic`
- KYC campaigns: `AR_Onboarding_Install_*` or `AR_Onboarding_OTP_*`
- TXN campaigns: `AR_Onboarding_Kyced_*`
- Retention: `AR_Retention_*`

### Wati
- Onboarding: `marketing_onboarding_*`
- Retention: `marketing_retention_*`
- KYC: `otp_entered_not_kyc` or `marketing_install`

### WebEngage In-app
- Subcategory promotions: `Subcategory_Promotion_{SubcatName}_{Seg}_{ScreenType}_{Date}`
- Tags must include `subccat` (note spelling)

---

## Notes

- **OneDrive sync**: Scripts write to local OneDrive folder. Sync to cloud happens automatically but shared web link may show cached version until file is opened in Excel desktop app.
- **Google Ads Unique Users**: Refreshes at scheduled report time (7 PM IST). Numbers before refresh reflect previous day.
- **WebEngage data**: Last 30 days rolling window. In-app subcategory data is accumulated in `Inapp_Subcat_Store` tab to prevent data loss across months.
- **MTD timing**: Scripts run after 12 PM IST use today as MTD end; before 12 PM use yesterday.