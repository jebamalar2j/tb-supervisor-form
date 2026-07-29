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
    d = (district or "").strip()
    t = (tu or "").strip()
    if d.lower() == "chennai":
        if t not in CHENNAI_1_TUS and t not in CHENNAI_2_TUS:
            print(f"    WARNING: unrecognized TU '{tu}' for Chennai - routing to Chennai_2 by default, please verify")
        return "Chennai_1" if t in CHENNAI_1_TUS else "Chennai_2"
    if not d:
        print("    WARNING: blank district on a submission - patient may be lost, check KoBo data")
    return d

def fetch_kobo_submissions():
    """Single paginated bulk pull of every submission on the server.
    Reused for both the patient sync (Sheet) and the daily test-count
    totals (email report) - only one API call series per run."""
    print("\nFetching all submissions from KoBoToolbox...")
    submissions = []
    url = f"{EU_API}/api/v2/assets/{LAB_FORM_ID}/data/?format=json&limit=5000"
    while url:
        resp = requests.get(url, headers=KOBO_HEADERS)
        if resp.status_code != 200:
            print(f"  ERROR: {resp.status_code}")
            break
        payload = resp.json()
        submissions.extend(payload.get("results", []))
        url = payload.get("next")  # paginate past 5000 if present
    print(f"  {len(submissions)} total submissions on server")
    return submissions


def extract_patients(submissions):
    """One row per unique Patient Ni-kshay ID, for the Google Sheet
    (Power BI merge only needs one NAAT-linkable row per patient - if
    the same patient shows up across multiple test sub-pages/days, only
    the first occurrence is kept here)."""
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

            row_date = rdate[:10] if rdate else ""

            if nid and str(nid) not in seen_ids:
                seen_ids.add(str(nid))
                patients.append({
                    "district":   rd,
                    "tu":         rtu,
                    "site":       rs,
                    "date":       row_date,
                    "nid":        str(nid),
                    "lab_serial": str(lab) if lab else "",
                    "mtb_result": str(mtb) if mtb else "",
                })

    print(f"  {len(patients)} unique patient(s) found")
    return patients


def extract_daily_test_counts(submissions):
    """Every row in group_mb5pc30, counted by its own 'Date of testing'.
    Each row is one individual-test sub-page (KoBo pops up N sub-pages
    when 'No. of Test done in the day' = N) - NOT related to the
    clinical 'Repeat Test Done' field, which is answered inside the same
    row/sub-page and does not create an extra row. A patient can
    legitimately appear on more than one day's sub-pages (re-presenting
    for testing later), and each such row is a real, separate test - so
    this is intentionally NOT deduplicated by patient ID. This is what
    the email report's 'Total tests' figures are built from."""
    counts = {}  # date_str -> count
    for sub in submissions:
        for row in sub.get("group_mb5pc30", []):
            rdate = str(row.get("group_mb5pc30/Date_of_testing", ""))
            nid = row.get("group_mb5pc30/Patient_Ni_kshay_ID")
            if not rdate or not nid:
                continue
            d = rdate[:10]
            counts[d] = counts.get(d, 0) + 1
    return counts



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
    """Set of Patient Ni-kshay IDs already in the sheet (column F). One
    row per patient is enough - the Sheet only needs to carry MTB Result
    + NAAT columns for the Power BI merge by Ni-kshay ID."""
    try:
        all_rows = ws.get_all_values()
        return {str(row[5]).strip() for row in all_rows[1:] if len(row) >= 6 and row[5]}
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

def process_date(gc, patients):
    if not patients:
        print("  No patients found")
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
            nid = str(p["nid"]).strip()
            if nid in existing:
                print(f"    Skipping {p['nid']} - already exists")
                continue
            existing.add(nid)
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
    if not value and value != 0:
        return None
    value = str(value).strip()
    if not value:
        return None

    formats = (
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
        "%d-%m-%Y", "%d-%b-%Y", "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y", "%Y/%m/%d",
    )
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    # Google Sheets serial date number (e.g. cell formatted as plain
    # number instead of a date string). Epoch is 1899-12-30.
    try:
        serial = float(value)
        if 20000 < serial < 80000:  # sane range for ~2000-2119
            from datetime import date as _date, timedelta as _td
            return _date(1899, 12, 30) + _td(days=serial)
    except ValueError:
        pass

    return None


def build_daily_report(gc, run_date, daily_counts):
    """Total tests today/till date come from daily_counts (raw KoBo
    'Date of testing' row counts, including repeats - see
    extract_daily_test_counts). Positive/NAAT/discordant/backlog stay
    Sheet-based (one row per unique patient)."""
    sh = gc.open_by_key(MASTER_SHEET_ID)
    cutoff = run_date - timedelta(days=REPORT_LAG_DAYS)

    total_tests_today = daily_counts.get(cutoff.strftime("%Y-%m-%d"), 0)
    total_tests_till_date = sum(
        n for d, n in daily_counts.items() if d <= cutoff.strftime("%Y-%m-%d")
    )

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

        skipped_no_date = 0
        skipped_no_id = 0
        bad_date_samples = []
        total_rows_seen = 0

        for r in values[1:]:
            def cell(idx):
                return r[idx].strip() if idx is not None and idx < len(r) else ""

            raw_date = cell(date_col)
            nid = cell(nid_col)
            if not raw_date and not nid:
                continue  # fully blank row, not worth counting

            total_rows_seen += 1
            test_date = parse_sheet_date(raw_date)

            if not nid:
                skipped_no_id += 1
                continue
            if not test_date:
                skipped_no_date += 1
                if len(bad_date_samples) < 5:
                    bad_date_samples.append(raw_date)
                continue

            rows.append({
                "nid": nid,
                "date": test_date,
                "mtb": cell(mtb_col),
                "cbnaat": cell(cbnaat_col).lower(),
                "truenat": cell(truenat_col).lower(),
            })

        print(
            f"  {tab_name}: {total_rows_seen} data rows, "
            f"{skipped_no_id} skipped (no Ni-kshay ID), "
            f"{skipped_no_date} skipped (unparseable date)"
        )
        if bad_date_samples:
            print(f"    sample unparseable date values: {bad_date_samples}")

    till_date_rows = [r for r in rows if r["date"] <= cutoff]

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
        f"Good morning Ma'am\n"
        f"TN UniAmp N-PoC \u2013 Daily Report\n"
        f"Data as on {m['cutoff'].strftime('%d-%b-%Y')}\n\n"
        f"Total tests today: {m['total_tests_today']}\n"
        f"Total tests till date: {m['total_tests_till_date']}\n"
        f"Total positive in UniAmp: {m['total_uniamp_positive']}\n"
        f"Total confirmed positive in NAAT: {m['total_naat_positive']}\n"
        f"Discordant: {m['discordant']}\n"
        f"Backlog to confirm in NAAT: {m['backlog']}\n\n"
        f"Thank you"
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

    print("\nPulling all KoBo submissions (every run is a full sync - "
          "no missed-night gaps possible)")

    gc = get_sheets_client()
    submissions = fetch_kobo_submissions()
    patients = extract_patients(submissions)
    daily_counts = extract_daily_test_counts(submissions)

    total = process_date(gc, patients)

    print(f"\n{'='*55}")
    print(f"OK Sync complete - {total} new row(s) added")
    print(f"  {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S IST')}")
    if total > 0:
        print(f"  https://docs.google.com/spreadsheets/d/{MASTER_SHEET_ID}")

    print(f"\n{'='*55}")
    print("  Daily report")
    print(f"{'='*55}")
    metrics = build_daily_report(gc, run_date=datetime.now(IST).date(), daily_counts=daily_counts)
    message = format_whatsapp_message(metrics)
    print(message)
    send_report_email(
        subject=f"TN UniAmp Daily Report - {metrics['run_date'].strftime('%d-%b-%Y')}",
        body=message,
    )

if __name__ == "__main__":
    main()
