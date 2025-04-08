import gspread
from google.oauth2.service_account import Credentials
import openai
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import markdown
from weasyprint import HTML
import io
from fpdf import FPDF
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import pymupdf
import pandas as pd
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file while maintaining structure."""
    doc = pymupdf.open(pdf_path)
    text = "\n".join([page.get_text("text") for page in doc])  # Extract text page by page
    return text

# pdf_path = "evaluation_report_karishma.shah@devoteam.com.pdf"
# pdf_content = extract_text_from_pdf(pdf_path)

date_format = "%d/%m/%Y %H:%M:%S"

# Get the current date and time
current_date = datetime.now()

# Google Sheets API Setup
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "parser_SA.json"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# OpenAI API Key
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")



def save_text_to_pdf(text, filename):
    # Create a PDF object
    pdf = FPDF()

    # Add a page
    pdf.add_page()

    # Set font (Arial, regular, 12pt)
    pdf.set_font('Arial', '', 12)

    # Insert the text into the PDF
    pdf.multi_cell(0, 10, text)

    # Output PDF to file
    pdf.output(filename)



def markdown_to_pdf(markdown_text: str, output_filename: str):
    """
    Converts a Markdown string into a PDF file with rendered formatting.
    
    :param markdown_text: A string containing Markdown content.
    :param output_filename: The name of the output PDF file.
    """
    # Convert Markdown to HTML
    html_content = markdown.markdown(markdown_text, extensions=['extra', 'tables'])
    
    # Wrap in a basic HTML template
    full_html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h2 {{ color: #333; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """
    
    # Convert HTML to PDF
    HTML(string=full_html).write_pdf(output_filename)
    print(f"PDF saved as {output_filename}")

# Authenticate with Google Sheets
def get_google_sheet():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    return sheet

# Fetch data from Google Sheets
# def fetch_employee_data():
#     sheet = get_google_sheet()
#     data = sheet.get_all_records()
#     return pd.DataFrame(data)

def fetch_employee_data():
    sheet = get_google_sheet()
    rows = sheet.get_all_values()  # Get all values as a list of lists

    # Ensure unique column names with suffixes (_1, _2, etc.)
    headers = rows[0]
    seen = {}
    unique_headers = []
    
    for col in headers:
        if col in seen:
            seen[col] += 1
            unique_headers.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique_headers.append(col)
    
    # Convert to DataFrame
    df = pd.DataFrame(rows[1:], columns=unique_headers)

    # Consolidate duplicate columns
    column_groups = {}
    for col in unique_headers:
        base_col = col.rsplit('_', 1)[0]  # Remove last "_X" to get base name
        column_groups.setdefault(base_col, []).append(col)

    for base_col, dup_cols in column_groups.items():
        if len(dup_cols) > 1:
            df[base_col] = df[dup_cols].bfill(axis=1).iloc[:, 0]  # Keep first non-null value
            df.drop(columns=dup_cols, inplace=True)  # Drop original duplicate columns

    df.to_csv("df_evaluation.csv", index= False)
    return df


# Format evaluation using OpenAI ChatGPT
def format_evaluation(evaluation_data):
    prompt = f"""
    Structure the following employee evaluation in a *tabular* format:
    {evaluation_data}
    """
    client = openai.OpenAI()
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "system", "content": "Format the following employee evaluation and the project details and client name, that is contained in a dictonary where the keys are the evaluation criteria and the value is the result of the evaluation, into a structured organized format to be presented in a document to his manager, make sure to to ALWAYS use proper markdown format and tabular display of information when appropriate, just ignore empty columns and columns not related to the evaluation criteria, and just mention in the title who's evlauation is it"},
                  {"role": "user", "content": prompt}]
    )

    
    return response.choices[0].message.content


# Send evaluation via email
def send_email(to_email, consultant_name, manager_email, flag_di=0):
    # html_content = markdown.markdown(evaluation_table)
    attachment_file = f"evaluation_report_{consultant_name}.pdf"
    # if os.path.isfile(attachment_file):
    #     print("email already sent!")
    #     return
    # # Convert HTML to PDF
    pdf_io = io.BytesIO()  # Create an in-memory file-like object
    # HTML(string=html_content).write_pdf(pdf_io)
    if flag_di:
        to_email_cc = ["saleh.samaneh@devoteam.com","mohamed.hatab@devoteam.com"]
    else:
        print("~~~~", to_email)
        to_email_cc = []
        
    # for to_email in to_email_l:
    
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = to_email
    msg['Subject'] = "Your Project Evaluation Form"
    msg['Cc'] = ', '.join(["dme.career.advisory@devoteam.com"] + to_email_cc + [manager_email])
    # Attach the body of the email
    msg.attach(MIMEText(f"Kindly find attached {consultant_name}'s performance evaluation for Q1 2025", 'plain'))

    attachment_file = f"evaluation_report_{consultant_name}.pdf"
    
    all_recipients = [to_email] + ["dme.career.advisory@devoteam.com"] + to_email_cc + [manager_email]  # Combine TO and CC for sending

    # Open the PDF file to be attached
    with open(attachment_file, "rb") as attachment:
        # Create a MIMEBase instance and encode the attachment
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)

        # Add header to specify the attachment filename
        part.add_header('Content-Disposition', f'attachment; filename={attachment_file}')

        # Attach the PDF to the email
        msg.attach(part)


    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, all_recipients, msg.as_string())

# Main Execution
if __name__ == "__main__":
    df = fetch_employee_data()
    
    for _, row in df.iterrows():

    # Convert the input string to a datetime object
        if row["Timestamp"]:
            print("~~~given date:", row["Timestamp"])
            given_date = datetime.strptime(row["Timestamp"], date_format)
            reference_time = datetime.strptime("26/03/2025 16:46:08", "%d/%m/%Y %H:%M:%S")
            if (given_date >= reference_time) and ('@' in row["Manager Email"]):
                evaluation_text = str(row.to_dict())  # Convert row to text 
                print("=============RAW=============")
                print(evaluation_text)
                print("=============RAW=============")
                formatted_evaluation = format_evaluation(evaluation_text)
                print("=============FORMATED=============")
                print(formatted_evaluation)
                print("=============FORMATED=============")
                markdown_to_pdf(formatted_evaluation, f"evaluation_report_{row['Consultant Email']}.pdf")
                flag_di = 0
                if "DI" in row['Business Unit']:
                    flag_di = 1
                
                send_email(row['Consultant Email'], row['Consultant Email'],row['Manager Email'], flag_di)
                print(f"email sent to: {row['Consultant Email']}")
    print("All emails sent successfully!")
