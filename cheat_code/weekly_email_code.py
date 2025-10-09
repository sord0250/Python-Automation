from pathlib import Path
import os
import smtplib, ssl
from email.message import EmailMessage
import mimetypes

import pandas as pd
from dotenv import load_dotenv

# This is the path to your current directory where your file is located
BASE_DIR = Path(__file__).resolve().parent

# Use your real filenames here
CSV_IN = BASE_DIR / "people.csv"        # <- name of your existing CSV
XLSX_OUT = BASE_DIR / "report.xlsx"   # <- name of excel export you will send in the email
# ----------------------------------------------------------

load_dotenv(BASE_DIR / ".env")  # SMTP settings from .env

# Env vars
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))  # wrap in an int so it is interpreted as an integer
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
EMAIL_FROM = os.getenv("EMAIL_FROM", "")


def csv_to_pretty_excel(csv_path: str, xlsx_path: str, sheet_name: str = "Data"):
    # Read CSV
    df = pd.read_csv(csv_path)

    # Write Excel with basic styling
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        # Header row (bold + light gray)
        header_fmt = wb.add_format({"bold": True, "bg_color": "#EEEEEE", "border": 1})
        for col_idx, name in enumerate(df.columns):
            ws.write(0, col_idx, name, header_fmt)

        # Auto-fit column widths (simple)
        for i, col in enumerate(df.columns):
            # max length between header and data (treat NaN as empty)
            series = df[col].astype(str).fillna("")
            max_len = max([len(col)] + series.map(len).tolist())
            ws.set_column(i, i, max_len + 2)

        # Calculate last_row and last_col (zero-based, inclusive)
        # Hint: header is at row 0, data starts at row 1
        last_row = len(df)                  # inclusive (0 is header, 1..len(df) are data rows)
        last_col = len(df.columns) - 1      # inclusive

        # Name your table, no spaces
        ws.add_table(0, 0, last_row, last_col, {
            "name": "DataTable",
            "columns": [{"header": c} for c in df.columns],
            "style": "Table Style Medium 9",
        })


def send_email(subject: str, body: str, attachment_path: str, cc=None, bcc=None):
    # Parse recipients from env (comma-separated)
    to_list = [e.strip() for e in os.getenv("EMAIL_TO", "").split(",") if e.strip()]
    cc_list = [] if not cc else ([cc] if isinstance(cc, str) else list(cc))
    bcc_list = [] if not bcc else ([bcc] if isinstance(bcc, str) else list(bcc))

    if not to_list:
        raise ValueError("EMAIL_TO is empty. Set EMAIL_TO in your .env (comma-separated if multiple).")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(body)

    # Attach the Excel file using mimetypes.guess_type and EmailMessage.add_attachment
    if attachment_path:
        mt, _ = mimetypes.guess_type(str(attachment_path))
        maintype, subtype = (mt.split("/", 1) if mt else ("application", "octet-stream"))
        with open(attachment_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=Path(attachment_path).name
            )

    # Send (STARTTLS flow; use 587)
    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
        # s.set_debuglevel(1)  # uncomment for SMTP conversation logs
        s.ehlo()
        s.starttls(context=context)
        s.ehlo()
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg, to_addrs=to_list + cc_list + bcc_list)


def main():
    # 1) Convert CSV to Excel
    csv_to_pretty_excel(CSV_IN, XLSX_OUT, sheet_name="Data")
    # 2) Email the Excel file
    send_email(
        subject="Automated Report",
        body="Hi! Please find the attached Excel report.",
        attachment_path=XLSX_OUT
    )


if __name__ == "__main__":
    main()
