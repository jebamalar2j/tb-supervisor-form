"""
ONE-TIME backfill: re-derives MTB Result for every row already in
Chennai_1 / Chennai_2, using the same G(mtb_result) + H(IC_Detected)
logic now used going forward, so old rows show "Invalid" where relevant
instead of blank/"Negative".

Run this ONCE (locally or via a manual GitHub Actions dispatch), then
delete/ignore it - the nightly tb_supervisor_sheets.py already applies
the correct derivation to all NEW rows.

Uses the same single bulk KoBo call pattern as the nightly sync (one
GET to /data.json with a high limit), matched to existing sheet rows by
Patient Ni-kshay ID. Does not create new rows or touch NAAT columns.
"""

import os
import sys
import requests
import gspread
from google.oauth2.service_account import Credentials

KOBO_API_KEY = os.environ.get("KOBO_API_KEY", "YOUR_KOBO_API_KEY_HERE")
LAB_FORM_ID = "aYPkk34YAhR6ZNJBhevCuQ"
EU_API = "https://eu.kobotoolbox.org"
MASTER_SHEET_ID = "1y5Jqv7wZG99uV9Eo0mGQuiEOEUsch-8odadj_Knt9l0"
REPORT_TABS = ["Chennai_1", "Chennai_2"]

KOBO_HEADERS = {"Authorization": f"Token {KOBO_API_KEY}"}


def get_sheets_client():
    sa_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_account.json")
    if not os.path.exists(sa_file):
        sa_file = "service_account.json"
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    return gspread.authorize(creds)


def derive_mtb_result(mtb_raw, ic_raw):
    if mtb_raw == "mtb_postive":
        return "MTB Positive"
    if mtb_raw == "negative" and ic_raw == "yes":
        return "Negative"
    if mtb_raw == "negative" and ic_raw == "no":
        return "Invalid"
    return ""


def fetch_all_derived_results():
    """One bulk call, all submissions, no date filter - covers every row
    ever synced so far."""
    url = f"{EU_API}/api/v2/assets/{LAB_FORM_ID}/data/?format=json&limit=5000"
    resp = requests.get(url, headers=KOBO_HEADERS)
    resp.raise_for_status()
    submissions = resp.json().get("results", [])
    print(f"  {len(submissions)} total submissions on server")

    derived_by_id = {}
    for sub in submissions:
        for row in sub.get("group_mb5pc30", []):
            nid = row.get("group_mb5pc30/Patient_Ni_kshay_ID")
            if not nid:
                continue
            mtb_raw = row.get("group_mb5pc30/mtb_result", "")
            ic_raw = row.get("group_mb5pc30/IC_Detected", "")
            derived_by_id[str(nid)] = derive_mtb_result(mtb_raw, ic_raw)
    return derived_by_id


def backfill():
    print("Fetching all KoBo submissions to re-derive MTB Result...")
    derived_by_id = fetch_all_derived_results()

    gc = get_sheets_client()
    sh = gc.open_by_key(MASTER_SHEET_ID)

    for tab_name in REPORT_TABS:
        ws = sh.worksheet(tab_name)
        header = ws.row_values(1)
        try:
            id_col = header.index("Patient Ni-kshay ID") + 1
            mtb_col = header.index("MTB Result") + 1
        except ValueError:
            print(f"  WARNING: expected columns not found in {tab_name}, skipping")
            continue

        ids = ws.col_values(id_col)[1:]  # skip header
        existing_mtb = ws.col_values(mtb_col)[1:]

        updates = []
        changed = 0
        for i, nid in enumerate(ids, start=2):  # actual sheet row number
            nid = str(nid).strip()
            new_val = derived_by_id.get(nid)
            old_val = existing_mtb[i - 2] if (i - 2) < len(existing_mtb) else ""
            if new_val and new_val != old_val:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, mtb_col),
                    "values": [[new_val]],
                })
                changed += 1

        if updates:
            ws.batch_update(updates)
        print(f"{tab_name}: {changed} row(s) corrected out of {len(ids)}")

    print("\nBackfill complete.")


if __name__ == "__main__":
    if KOBO_API_KEY == "YOUR_KOBO_API_KEY_HERE":
        print("ERROR: set KOBO_API_KEY environment variable first.")
        sys.exit(1)
    backfill()
