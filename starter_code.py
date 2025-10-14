from pathlib import Path
import os
import smtplib, ssl
from email.message import EmailMessage
import mimetypes

import pandas as pd
from dotenv import load_dotenv

# This is that path to your current directory where your file is located
BASE_DIR = Path(__file__).resolve().parent

# TO-DO: make sure it reads in the name of the correct csv file. Also name the exported report, make sure it is .xlsx
CSV_IN = BASE_DIR / " "      # <- name of your existing CSV
XLSX_OUT = BASE_DIR / " "    # <- name of excel export you will send in the email
# ----------------------------------------------------------

load_dotenv(BASE_DIR / ".env")  # Loads in the .env file so it can use the variables

# TO-DO: following the example of the first, read in the rest of the .env values
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = # wrap in an int so it is interpreted as an integer
SMTP_USER = 
SMTP_PASS = 

EMAIL_FROM = os.getenv("EMAIL_FROM", "") # don't change this one


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

        # TO-DO: using the python len() function, figure out how to calculate last_row and last_col
        # Hint: header is at row 0, data starts at row 1
        last_row = 
        last_col = 

       # TO-DO: Name your table, no spaces
        ws.add_table(0, 0, last_row, last_col, {
            "name": " ",
            "columns": [{"header": c} for c in df.columns],
            "style": "Table Style Medium 9",
        })

def send_email(subject:str, body:str, attachment_path:str):
    # Parse recipients from env
    to_list = [e.strip() for e in os.getenv("EMAIL_TO", "").split(",") if e.strip()]

    # Prepares the email's information
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(to_list)
    
    msg.set_content(body)

    # TO-DO: create code that will attach the excel file using the attatchment_path
    # Use mimetypes.guess_type and EmailMessage.add_attachment
    # This is if you are up for a challenge - depending on level, just copy this part from the cheat_code folder in the Github 
    if attachment_path:
        pass

    # send
    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
        # s.set_debuglevel(1) # uncomment if you want to see debug and active connection results
        # TO-DO: use s, .ehlo(), .starttls(), .ehlo(), the .login function and .send_message() from the smtp lib to login to the account
        # s is your connection to SMTP


def main():
    # TO-DO: put together your functions above in a simple main function that will:
    # 1) first take the csv and turn it into an excel 
    # 2) and then send it in an email with a subject and a body


if __name__ == "__main__":
    main()
