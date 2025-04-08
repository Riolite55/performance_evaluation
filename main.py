import gspread
from google.oauth2.service_account import Credentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import markdown
import io
from fpdf import FPDF
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import fitz as pymupdf
import pandas as pd
import os
from dotenv import load_dotenv
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.units import inch
import markdown2

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



def format_data(data):
    """
    Formats the raw form data into a markdown string.
    """
    timestamp = data.get('Timestamp', '')
    email = data.get('Email address', '')
    dme_id = data.get('DME ID - Employee Name', '')
    
    markdown_text = f"# Performance Evaluation Report\n\n"
    
    # Basic Information Table
    markdown_text += "## Basic Information\n\n"
    markdown_text += "<table>\n"
    markdown_text += "<tr><th>Field</th><th>Value</th></tr>\n"
    markdown_text += f"<tr><td>DME ID</td><td>{dme_id}</td></tr>\n"
    markdown_text += f"<tr><td>Email</td><td>{email}</td></tr>\n"
    markdown_text += f"<tr><td>Business Unit</td><td>{data.get('Business Unit', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Grade</td><td>{data.get('Employee Grade', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Manager</td><td>{data.get('Manager', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Manager Email</td><td>{data.get('Manager Email', '')}</td></tr>\n"
    markdown_text += "</table>\n\n"
    
    # Role and Profile Information
    markdown_text += "## Role & Technical Profile\n\n"
    markdown_text += "<table>\n"
    markdown_text += "<tr><th>Category</th><th>Details</th></tr>\n"
    markdown_text += f"<tr><td>Profile</td><td>{data.get('Profile assigned on the project', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Technical Role</td><td>{data.get('Technical Role Played on the Project', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Technical Capability</td><td>{data.get('Technical Capability Utilized on the Project', '')}</td></tr>\n"
    markdown_text += f"<tr><td>Quarter</td><td>{data.get('Evaluation filled for which quarter?', '')}</td></tr>\n"
    markdown_text += "</table>\n\n"
    
    # Project Performance
    if data.get('Project Name 3', ''):
        markdown_text += "## Project Performance: TVTC-PMO\n\n"
        markdown_text += "<table>\n"
        markdown_text += "<tr><th>Evaluation Criteria</th><th>Rating</th><th>Status</th></tr>\n"
        
        def get_status_html(rating):
            if rating == 'Exceeds Requirements':
                return '<span class="status-icon status-exceeds">★ Exceeds</span>'
            elif rating == 'Meets Requirements':
                return '<span class="status-icon status-meets">✓ Meets</span>'
            elif rating == 'Below Requirements':
                return '<span class="status-icon status-below">! Below</span>'
            return ''
        
        project_rating = data.get('Project/Deliverables Sign-off', '')
        feedback_rating = data.get('Providing Resources With Timely Constructive Feedback', '')
        
        markdown_text += f"<tr><td>Project/Deliverables Sign-off</td><td>{project_rating}</td><td>{get_status_html(project_rating)}</td></tr>\n"
        markdown_text += f"<tr><td>Timely Resource Feedback</td><td>{feedback_rating}</td><td>{get_status_html(feedback_rating)}</td></tr>\n"
        markdown_text += "</table>\n\n"
    
    # Behavioral Competencies
    markdown_text += "## Behavioral Competencies\n\n"
    markdown_text += "<table>\n"
    markdown_text += "<tr><th>Competency</th><th>Rating</th><th>Status</th></tr>\n"
    
    competencies = [
        ('Devoteam Values', 'Behavioral Competencies in accordance to Devoteam Values'),
        ('Trusted Mindset', 'Behavioral Competencies in accordance to Trusted Deavoteamer\'s Mindset'),
        ('Knowledge & Offerings', 'Knowledge of Devoteam roles, M0 & service offerings and their application in the current assigned client environment'),
        ('Feedback Response', 'Responsiveness to constructive feedback'),
        ('Collaboration', 'Collaboration & effective knowledge sharing')
    ]
    
    for display_name, key in competencies:
        rating = data.get(key, '')
        markdown_text += f"<tr><td>{display_name}</td><td>{rating}</td><td>{get_status_html(rating)}</td></tr>\n"
    
    markdown_text += "</table>\n\n"
    
    # Performance Recommendations
    markdown_text += "## Performance Improvement Recommendations\n\n"
    recommendations = data.get('Based on the assessment, describe how the employee can elevate their performance to deliver better outcomes and achieve greater client satisfaction during the project assignment.', '')
    if recommendations:
        markdown_text += f"<ul><li>{recommendations}</li></ul>\n\n"
    
    markdown_text += f"<p><em>Evaluation Date: {timestamp}</em></p>"
    
    return markdown_text

def markdown_to_pdf(markdown_text: str, output_filename: str):
    """
    Converts a Markdown string into a PDF file with rendered formatting using xhtml2pdf.
    
    Args:
        markdown_text (str): A string containing Markdown content
        output_filename (str): The name of the output PDF file
    """
    # Convert Markdown to HTML
    html_content = markdown2.markdown(markdown_text, extras=['tables', 'fenced-code-blocks'])
    
    # Create a complete HTML document with styling
    html_template = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.6;
                margin: 40px;
                color: #2c3e50;
            }}
            h1 {{
                font-size: 24px;
                text-align: center;
                margin-bottom: 30px;
            }}
            h2 {{
                font-size: 18px;
                margin-top: 20px;
                margin-bottom: 10px;
                color: #2c3e50;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 10px 0;
            }}
            th, td {{
                border: 1px solid #dee2e6;
                padding: 8px;
                text-align: left;
            }}
            th {{
                background-color: #f8f9fa;
                font-weight: bold;
            }}
            .status-icon {{
                font-weight: bold;
                font-size: 14px;
            }}
            .status-exceeds {{
                color: #2ecc71;
            }}
            .status-meets {{
                color: #3498db;
            }}
            .status-below {{
                color: #e74c3c;
            }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """
    
    # Convert HTML to PDF using xhtml2pdf
    from xhtml2pdf import pisa
    
    with open(output_filename, "w+b") as output_file:
        pisa.CreatePDF(html_template, dest=output_file)
    
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


def format_evaluation_deterministic(evaluation_data):
    """
    Format evaluation data in a deterministic way, creating a structured markdown document.
    
    Args:
        evaluation_data (dict): Dictionary containing evaluation data
        
    Returns:
        str: Formatted markdown string
    """
    # Define behavioral competency keys to exclude from project criteria
    behavioral_keys = {
        "Behavioral Competencies in accordance to Devoteam Values",
        "Behavioral Competencies in accordance to Trusted Deavoteamer's Mindset",
        "Knowledge of Devoteam roles, M0 & service offerings and their application in the current assigned client environment",
        "Responsiveness to constructive feedback",
        "Collaboration & effective knowledge sharing"
    }
    
    # Start building the markdown
    markdown_content = []
    
    # Add title with consultant name
    consultant_name = evaluation_data.get("DME ID - Employee Name", "Unknown Consultant")
    markdown_content.append(f"# Performance Evaluation: {consultant_name}\n")
    
    # Basic Information Section
    markdown_content.append("## Basic Information")
    basic_info = [
        ("Email Address", "Email address"),
        ("DME ID", "DME ID - Employee Name"),
        ("Business Unit", "Business Unit"),
        ("Employee Grade", "Employee Grade"),
        ("Profile", "Profile assigned on the project"),
        ("Technical Role", "Technical Role Played on the Project"),
        ("Technical Capability", "Technical Capability Utilized on the Project"),
        ("Evaluation Quarter", "Evaluation filled for which quarter?"),
        ("Manager", "Manager"),
        ("Manager Email", "Manager Email")
    ]
    
    for display_name, key in basic_info:
        value = evaluation_data.get(key)
        if value and str(value).strip() and str(value).lower() not in ['n/a', 'na', 'none', '', '0']:
            markdown_content.append(f"- **{display_name}**: {value}")
    
    markdown_content.append("")
    
    # Process each project
    for project_num in range(1, 8):
        project_name_key = f"Project Name {' ' if project_num == 3 else ''}{project_num}"
        project_name = evaluation_data.get(project_name_key)
        
        if project_name and str(project_name).strip() and str(project_name).lower() not in ['n/a', 'na', 'none', '', '0']:
            markdown_content.append(f"## Project {project_num}: {project_name}")
            
            # Project details
            client_name_key = f"Client Name{f' {project_num}' if project_num > 1 else ''}"
            client_name = evaluation_data.get(client_name_key)
            if client_name and str(client_name).strip() and str(client_name).lower() not in ['n/a', 'na', 'none', '', '0']:
                markdown_content.append(f"- **Client**: {client_name}")
            
            project_details = [
                ("Assignment Date", f"Project assignment date{f' {project_num}' if project_num > 1 else ''}"),
                ("Start Date", f"Project start date{f' {project_num}' if project_num > 1 else ''}"),
                ("DRM Name", "DRM Name" if project_num == 1 else None),
                ("BDM Name", "BDM Name" if project_num == 1 else None),
                ("CRP/CRD", f"CRP/CRD - Client Relationship Partner/Director{f' {project_num}' if project_num > 1 else ''}")
            ]
            
            for display_name, key in project_details:
                if key:  # Skip if key is None
                    value = evaluation_data.get(key)
                    if value and str(value).strip() and str(value).lower() not in ['n/a', 'na', 'none', '', '0']:
                        markdown_content.append(f"- **{display_name}**: {value}")
            
            # Project evaluation criteria
            project_criteria = []
            for key, value in evaluation_data.items():
                if (
                    value and str(value).strip() and str(value).lower() not in ['n/a', 'na', 'none', '', '0']
                    and not key.startswith(("Timestamp", "Email", "DME ID", "Business Unit", "Employee Grade", "Profile", "Technical Role", "Evaluation", "Manager", "Project Name", "Client Name", "Project assignment", "Project start", "DRM Name", "BDM Name", "CRP/CRD", "Based on the assessment"))
                    and key not in behavioral_keys
                ):
                    # Format the display name
                    display_name = key.split(" - ")[0] if " - " in key else key
                    project_criteria.append((display_name, value))
            
            if project_criteria:
                markdown_content.append("\n### Evaluation Criteria")
                for name, value in project_criteria:
                    markdown_content.append(f"- **{name}**: {value}")
            
            markdown_content.append("")
    
    # Behavioral Competencies Section
    behavioral_data = [(key, evaluation_data.get(key)) for key in behavioral_keys 
                      if evaluation_data.get(key) and str(evaluation_data.get(key)).strip() 
                      and str(evaluation_data.get(key)).lower() not in ['n/a', 'na', 'none', '', '0']]
    
    if behavioral_data:
        markdown_content.append("## Behavioral Competencies")
        for key, value in behavioral_data:
            display_name = key.split(" in accordance to ")[-1] if " in accordance to " in key else key
            markdown_content.append(f"- **{display_name}**: {value}")
        markdown_content.append("")
    
    # Performance Improvement Section
    improvement = evaluation_data.get("Based on the assessment, describe how the employee can elevate their performance to deliver better outcomes and achieve greater client satisfaction during the project assignment.")
    if improvement and str(improvement).strip() and str(improvement).lower() not in ['n/a', 'na', 'none', '', '0']:
        markdown_content.append("## Performance Improvement Recommendations")
        markdown_content.append(improvement)
        markdown_content.append("")
    
    # Add timestamp if available
    if "Timestamp" in evaluation_data:
        markdown_content.append(f"\n*Evaluation Date: {evaluation_data['Timestamp']}*")
    
    return "\n".join(markdown_content)

# Replace the old format_evaluation function with the new one
def format_evaluation(evaluation_data):
    """
    Format evaluation data into a structured markdown document.
    
    Args:
        evaluation_data: Can be either a string representation of a dictionary or a dictionary
        
    Returns:
        str: Formatted markdown string
    """
    # If input is a string, convert it to a dictionary
    if isinstance(evaluation_data, str):
        # Remove any leading/trailing whitespace and convert single quotes to double quotes for proper JSON parsing
        cleaned_data = evaluation_data.strip().replace("'", '"')
        try:
            import json
            evaluation_dict = json.loads(cleaned_data)
        except json.JSONDecodeError:
            # If JSON parsing fails, try using ast.literal_eval as a fallback
            import ast
            try:
                evaluation_dict = ast.literal_eval(evaluation_data)
            except (ValueError, SyntaxError):
                raise ValueError("Invalid evaluation data format")
    else:
        evaluation_dict = evaluation_data
    
    return format_evaluation_deterministic(evaluation_dict)

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
            reference_time = datetime.strptime("21/03/2025 10:46:11", "%d/%m/%Y %H:%M:%S")
            if (given_date == reference_time) and ('@' in row["Manager Email"]):
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
