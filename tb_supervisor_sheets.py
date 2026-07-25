"""
TB Supervisor - Google Sheets Auto-Populator v5
================================================
- Runs automatically every night at 2 AM IST
- Pulls ALL districts/sites for previous day
- Routes to correct tab automatically
- Handles UTC to IST date conversion
"""

import requests
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import os
import sys
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

KOBO_API_KEY    = os.environ.get("KOBO_API_KEY", "YOUR_KOBO_API_KEY_HERE")
LAB_FORM_ID     = "aYPkk34YAhR6ZNJBhevCuQ"
EU_API          = "https://eu.kobotoolbox.org"
MASTER_SHEET_ID = "1y5Jqv7wZG99uV9Eo0mGQuiEOEUsch-8odadj_Knt9l0"

CHENNAI_1_TUS = [
    "Thiruvetriyur","Madhavaram","Tondiarpet","Kodungaiyur",
    "Basin_Bridge","Kolathur","Padi_Manjakuppam","Aminjikarai",
    "Mmda","Mylapore","Virugambakkam"
]
CHENNAI_2_TUS = [
    "Kodambakkam_Tu","Saidapet_West","Chinnaporur","Nanganallur",
    "Velachery","Puzhithivakkam","Kannagi_Nagar"
]

HEADERS = [
    "Date of Testing", "District", "Name of TU", "Name of Site / DMC",
    "Lab Serial No", "Patient Ni-kshay ID", "MTB Result",
    "CBNAAT Result", "Truenat testing status", "X-ray testing status",
    "Notification", "Rif resistance testing status", "Other resistance",
    "Treatment initiated", "Type of treatment initiated",
    "TruNAAT CFU value", "CBNAAT result", "Remarks"
]

# Dropdowns on H(7) I(8) J(9) K(10) L(11) N(13) O(14) Q(16) only
# G(6) = MTB Result - NO dropdown, pre-filled from KoBoToolbox
DROPDOWN_COLS = {
    7:  ["MTB Detected", "MTB Not detected"],
    8:  ["MTB Detected", "MTB Not detected"],
    9:  ["Suggestive of TB", "Not suggestive of TB"],
    10: ["Microbiologically confirmed TB", "Clinically diagnosed TB",
         "Follow-up case", "Not a case"],
    11: ["Resistant", "Not resistant", "Other resistance", "Indeterminate"],
    13: ["Yes", "Not required"],
    14: ["DS-TB", "DR-TB"],
    16: ["Very low", "Low", "Medium", "High"],
}

KOBO_HEADERS = {"Authorization": f"Token {KOBO_API_KEY}"}
IST = ZoneInfo("Asia/Kolkata")

# --- Daily report config -----------------------------------------------
REPORT_TABS = ["Chennai_1", "Chennai_2"]  # only active tabs; add district
                                           # tabs here once they start testing
REPORT_LAG_DAYS = 2  # sites enter data same day / next day / day after
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

def get_sheets_client():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("ERROR: Run: pip install gspread google-auth")
        sys.exit(1)
    sa_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_account.json")
    if not os.path.exists(sa_file):
        sa_file = os.path.join(os.getcwd(), "service_account.json")
    if not os.path.exists(sa_file):
        sa_file = "service_account.json"
    if not os.path.exists(sa_file):
        print("ERROR: service_account.json not found.")
        sys.exit(1)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    from google.oauth2.service_account import Credentials
    import gspread
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    return gspread.authorize(creds)

def get_tab_name(district, tu):
    if district == "Chennai":
        return "Chennai_1" if tu in CHENNAI_1_TUS else "Chennai_2"
    return district

def fetch_all_patients(target_date):
    """Fetch ALL patients for a given IST date."""
    print(f"\nFetching all patients for {target_date} (IST) from KoBoToolbox...")
    url = f"{EU_API}/api/v2/assets/{LAB_FORM_ID}/data/?format=json&limit=5000"
    resp = requests.get(url, headers=KOBO_HEADERS)
    if resp.status_code != 200:
        print(f"  ERROR: {resp.status_code}")
        return []
    submissions = resp.json().get("results", [])
    print(f"  {len(submissions)} total submissions on server")

    patients = []
    seen_ids = set()

    for sub in submissions:
        for row in sub.get("group_mb5pc30", []):
            rd    = row.get("group_mb5pc30/district_repeat", "")
            rs    = row.get("group_mb5pc30/facility_name_repeat", "")
            rtu   = row.get("group_mb5pc30/tu_repeat", "")
            rdate = str(row.get("group_mb5pc30/Date_of_testing", ""))
            nid   = row.get("group_mb5pc30/Patient_Ni_kshay_ID")
            lab   = row.get("group_mb5pc30/Lab_serial_no", "")
            mtb_raw = row.get("group_mb5pc30/mtb_result", "")
            ic_raw  = row.get("group_mb5pc30/IC_Detected", "")

            # Derive MTB Result from mtb_result (G) + IC_Detected (H):
            #   mtb_postive              -> MTB Positive
            #   negative + IC Detected=yes -> Negative
            #   negative + IC Detected=no  -> Invalid
            if mtb_raw == "mtb_postive":
                mtb = "MTB Positive"
            elif mtb_raw == "negative" and ic_raw == "yes":
                mtb = "Negative"
            elif mtb_raw == "negative" and ic_raw == "no":
                mtb = "Invalid"
            else:
                mtb = ""  # unexpected/blank combination - don't guess

            # Match date - KoBoToolbox stores date in YYYY-MM-DD (IST)
            if rdate.startswith(target_date) and nid and str(nid) not in seen_ids:
                seen_ids.add(str(nid))
                patients.append({
                    "district":   rd,
                    "tu":         rtu,
                    "site":       rs,
                    "date":       target_date,
                    "nid":        str(nid),
                    "lab_serial": str(lab) if lab else "",
                    "mtb_result": str(mtb) if mtb else "",
                })

    print(f"  {len(patients)} total patient(s) found")
    return patients

def get_or_create_tab(sh, tab_name):
    try:
        ws = sh.worksheet(tab_name)
    except Exception:
        print(f"  Creating new tab: {tab_name}")
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=18)
    first_row = ws.row_values(1)
    if not first_row or first_row[0] != "Date of Testing":
        ws.insert_row(HEADERS, 1)
        print(f"  Headers added to {tab_name}")
    return ws

def get_existing_ids(ws):
    try:
        all_rows = ws.get_all_values()
        return {str(row[5]) for row in all_rows[1:] if len(row) >= 6 and row[5]}
    except Exception:
        return set()

def add_dropdowns_to_rows(sh, ws, start_row, end_row):
    if start_row > end_row:
        return
    sheet_id = ws._properties["sheetId"]
    requests_body = []
    for col_idx, options in DROPDOWN_COLS.items():
        requests_body.append({
            "setDataValidation": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    start_row - 1,
                    "endRowIndex":      end_row,
                    "startColumnIndex": col_idx,
                    "endColumnIndex":   col_idx + 1,
                },
                "rule": {
                    "condition": {
                        "type":   "ONE_OF_LIST",
                        "values": [{"userEnteredValue": o} for o in options],
                    },
                    "showCustomUi": True,
                    "strict":       True,
                }
            }
        })
    if requests_body:
        sh.batch_update({"requests": requests_body})
        print(f"    OK Dropdowns added to rows {start_row}-{end_row}")

def process_date(gc, target_date):
    patients = fetch_all_patients(target_date)
    if not patients:
        print(f"  No patients found for {target_date}")
        return 0

    sh = gc.open_by_key(MASTER_SHEET_ID)

    # Group by tab
    by_tab = {}
    for p in patients:
        tab = get_tab_name(p["district"], p["tu"])
        by_tab.setdefault(tab, []).append(p)

    total_added = 0
    for tab_name, tab_patients in by_tab.items():
        print(f"\n  Tab: {tab_name} - {len(tab_patients)} patient(s)")
        ws       = get_or_create_tab(sh, tab_name)
        existing = get_existing_ids(ws)

        new_rows = []
        for p in tab_patients:
            if p["nid"] in existing:
                print(f"    Skipping {p['nid']} - already exists")
                continue
            new_rows.append([
                p["date"], p["district"], p["tu"], p["site"],
                p["lab_serial"], p["nid"], p["mtb_result"],
                "", "", "", "", "", "", "", "", "", "", ""
            ])

        if not new_rows:
            print(f"    No new rows for {tab_name}")
            continue

        all_values = ws.get_all_values()
        start_row  = len(all_values) + 1
        end_row    = start_row + len(new_rows) - 1

        ws.append_rows(new_rows, value_input_option="RAW")
        add_dropdowns_to_rows(sh, ws, start_row, end_row)

        for row in new_rows:
            print(f"    OK {row[5]} | Lab: {row[4]} | MTB: {row[6]}")

        total_added += len(new_rows)

    return total_added

def find_col(header, keywords):
    """Find a column index (0-based) whose header contains any keyword,
    case-insensitive. Used because header text has drifted slightly from
    the HEADERS constant above (e.g. 'CBNAAT Result' vs 'CBNAAT testing
    status') without breaking the positional writes above."""
    for i, h in enumerate(header):
        h_low = (h or "").strip().lower()
        for kw in keywords:
            if kw in h_low:
                return i
    return None


def parse_sheet_date(value):
    if not value:
        return None
    value = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def build_daily_report(gc, run_date):
    """Reads the already-synced Sheet tabs (no extra KoBo calls) and
    computes the 6 daily metrics."""
    sh = gc.open_by_key(MASTER_SHEET_ID)
    cutoff = run_date - timedelta(days=REPORT_LAG_DAYS)

    rows = []
    for tab_name in REPORT_TABS:
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            continue
        values = ws.get_all_values()
        if not values:
            continue
        header = [h.strip() for h in values[0]]

        date_col   = find_col(header, ["date of testing"])
        nid_col    = find_col(header, ["ni-kshay", "nikshay", "ni_kshay"])
        mtb_col    = find_col(header, ["mtb result"])
        cbnaat_col = find_col(header, ["cbnaat"])
        truenat_col = find_col(header, ["truenat"])

        if None in (date_col, nid_col, mtb_col):
            print(f"  WARNING: could not find expected columns in {tab_name}, skipping for report")
            continue

        for r in values[1:]:
            def cell(idx):
                return r[idx].strip() if idx is not None and idx < len(r) else ""

            test_date = parse_sheet_date(cell(date_col))
            nid = cell(nid_col)
            if not test_date or not nid:
                continue

            rows.append({
                "nid": nid,
                "date": test_date,
                "mtb": cell(mtb_col),
                "cbnaat": cell(cbnaat_col).lower(),
                "truenat": cell(truenat_col).lower(),
            })

    till_date_rows = [r for r in rows if r["date"] <= cutoff]
    today_rows = [r for r in rows if r["date"] == cutoff]

    total_tests_today = len(today_rows)
    total_tests_till_date = len(till_date_rows)
    total_uniamp_positive = sum(1 for r in till_date_rows if r["mtb"] == "MTB Positive")

    def naat_status(r):
        if r["cbnaat"] == "mtb detected" or r["truenat"] == "mtb detected":
            return "Positive"
        if r["cbnaat"] == "mtb not detected" or r["truenat"] == "mtb not detected":
            return "Negative"
        return ""

    total_naat_positive = sum(1 for r in rows if naat_status(r) == "Positive")

    discordant = 0
    backlog = 0
    for r in till_date_rows:
        naat = naat_status(r)
        if r["mtb"] == "MTB Positive":
            if naat == "Negative":
                discordant += 1
            elif naat == "":
                backlog += 1
        elif r["mtb"] == "Negative" and naat == "Positive":
            discordant += 1

    return {
        "run_date": run_date,
        "cutoff": cutoff,
        "total_tests_today": total_tests_today,
        "total_tests_till_date": total_tests_till_date,
        "total_uniamp_positive": total_uniamp_positive,
        "total_naat_positive": total_naat_positive,
        "discordant": discordant,
        "backlog": backlog,
    }


def format_whatsapp_message(m):
    return (
        f"*TN UniAmp N-PoC \u2013 Daily Report*\n"
        f"_Data as on {m['cutoff'].strftime('%d-%b-%Y')} (2-day lag applied)_\n\n"
        f"Total tests today: *{m['total_tests_today']}*\n"
        f"Total tests till date: *{m['total_tests_till_date']}*\n"
        f"Total positive in UniAmp: *{m['total_uniamp_positive']}*\n"
        f"Total confirmed positive in NAAT: *{m['total_naat_positive']}*\n"
        f"Discordant: *{m['discordant']}*\n"
        f"Backlog to confirm in NAAT: *{m['backlog']}*"
    )


def send_report_email(subject, body):
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        print("  WARNING: GMAIL_ADDRESS/GMAIL_APP_PASSWORD not set - skipping email")
        return
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = GMAIL_ADDRESS
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.send_message(msg)
    print("  OK Report emailed")


def main():
    print("=" * 55)
    print("  TB Supervisor - Google Sheets Populator v5")
    print("=" * 55)

    if KOBO_API_KEY == "YOUR_KOBO_API_KEY_HERE":
        print("\nERROR: Please set your KoBoToolbox API key.")
        return

    # Use argument date if provided, otherwise yesterday IST
    if len(sys.argv) > 1:
        target_date = sys.argv[1]
        print(f"\nUsing specified date: {target_date}")
    else:
        # At 2 AM IST we want previous day's data
        yesterday_ist = datetime.now(IST) - timedelta(days=1)
        target_date   = yesterday_ist.strftime("%Y-%m-%d")
        print(f"\nAuto mode - pulling yesterday's data: {target_date} (IST)")

    gc = get_sheets_client()
    total = process_date(gc, target_date)

    print(f"\n{'='*55}")
    print(f"OK Sync complete - {total} new row(s) added")
    print(f"  {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S IST')}")
    if total > 0:
        print(f"  https://docs.google.com/spreadsheets/d/{MASTER_SHEET_ID}")

    print(f"\n{'='*55}")
    print("  Daily report")
    print(f"{'='*55}")
    metrics = build_daily_report(gc, run_date=datetime.now(IST).date())
    message = format_whatsapp_message(metrics)
    print(message)
    send_report_email(
        subject=f"TN UniAmp Daily Report - {metrics['run_date'].strftime('%d-%b-%Y')}",
        body=message,
    )

if __name__ == "__main__":
    main()
