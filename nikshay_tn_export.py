"""
Nikshay Tamil Nadu export - WEEKLY, Google Sheets only.

Reads the TB_Supervisor_Follow-up_Data sheet (Chennai_1, Chennai_2 -
add more tabs to REPORT_TABS once other districts start testing),
reorders/renames columns to match the Nikshay Maharashtra layout
exactly, and emails the result as an .xlsx attachment with a sheet
named "Nikshay Tamil Nadu".

Does NOT call the KoBoToolbox API - runs independently of the nightly
sync workflow, on its own weekly schedule.
"""

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

MASTER_SHEET_ID = "1y5Jqv7wZG99uV9Eo0mGQuiEOEUsch-8odadj_Knt9l0"
REPORT_TABS = ["Chennai_1", "Chennai_2"]  # add district tabs here once active

GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
IST = ZoneInfo("Asia/Kolkata")

# (target_column_name, source_TN_header_or_None)
# Headers match the "Nikshay Maharashtra" tab exactly, in the same
# order, so the two files can be appended directly in Power BI. Where
# TN's form doesn't collect that field, the header is kept and every
# row is left blank for that column.
NIKSHAY_EXPORT_COLUMNS = [
    ("district_repeat", "District"),
    ("tu_repeat", "Name of TU"),
    ("facility_name_repeat", "Name of Site / DMC"),
    ("total_no_of_test_repeat", None),
    ("test_no", None),
    ("Test ${test_no}", None),
    ("Patient Ni-kshay ID", "Patient Ni-kshay ID"),
    ("MTB Result", "MTB Result"),
    ("Microscopy testing status", None),
    ("CBNAAT testing status", "CBNAAT Result"),
    ("Truenat testing status", "Truenat testing status"),
    ("X-ray testing status", "X-ray testing status"),
    ("Notification", "Notification"),
    ("Rif resistance testing status", "Rif resistance testing status"),
    ("Treatment initiated", "Treatment initiated"),
    ("Type of treatment initiated DS-TB/DR-TB", "Type of treatment initiated"),
    ("TruNAAT CFU value", "TruNAAT CFU value"),
    ("CBNAAT result", "CBNAAT result"),
    ("Repeat Test Done", None),
    ("MTB Positive Repeat Result", None),
    ("Repeat Sputum sample sent for NAAT Confirmation", None),
    ("Total No. of positive repeat samples sent for NAAT confirmation", None),
    ("Lab serial no", "Lab Serial No"),
    ("Patient Ni-kshay ID2", None),
    ("Date of testing", "Date of Testing"),
    ("Patient Type", None),
    ("If Others, Please specify Patient Type", None),
    ("Visual appearance of sputum specimen", None),
    ("Sample transported for confirmatory  NAAT", None),
    ("Total No. of positive samples sent for NAAT confirmation", None),
    ("Remarks", "Remarks"),
]


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


def build_nikshay_tn_export(gc, out_path="Nikshay_Tamil_Nadu.xlsx"):
    sh = gc.open_by_key(MASTER_SHEET_ID)
    all_rows = []

    for tab_name in REPORT_TABS:
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            continue
        values = ws.get_all_values()
        if not values:
            continue
        header = [h.strip() for h in values[0]]
        col_index = {h: i for i, h in enumerate(header)}

        for r in values[1:]:
            def get(src_name):
                if src_name is None or src_name not in col_index:
                    return ""
                idx = col_index[src_name]
                return r[idx] if idx < len(r) else ""

            if not get("Patient Ni-kshay ID"):
                continue  # skip blank rows

            row_out = [get(src_name) for _target, src_name in NIKSHAY_EXPORT_COLUMNS]
            all_rows.append(row_out)

    wb = Workbook()
    ws_out = wb.active
    ws_out.title = "Nikshay Tamil Nadu"
    ws_out.append([target for target, _src in NIKSHAY_EXPORT_COLUMNS])
    for row in all_rows:
        ws_out.append(row)
    wb.save(out_path)
    print(f"OK Nikshay Tamil Nadu export written: {out_path} ({len(all_rows)} rows)")
    return out_path


def send_export_email(subject, body, attachment_path):
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        print("WARNING: GMAIL_ADDRESS/GMAIL_APP_PASSWORD not set - skipping email")
        return

    msg = MIMEMultipart()
    msg.attach(MIMEText(body))
    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{os.path.basename(attachment_path)}"',
    )
    msg.attach(part)
    msg["Subject"] = subject
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = GMAIL_ADDRESS

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.send_message(msg)
    print("OK Export emailed")


def main():
    gc = get_sheets_client()
    export_path = build_nikshay_tn_export(gc)
    run_date = datetime.now(IST).strftime("%d-%b-%Y")
    send_export_email(
        subject=f"Nikshay Tamil Nadu Export - {run_date}",
        body="Weekly Nikshay Tamil Nadu export attached.",
        attachment_path=export_path,
    )


if __name__ == "__main__":
    main()
