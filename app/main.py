from fastapi import FastAPI, HTTPException, Header
from fastapi import Body
import gspread
from google.auth import exceptions
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from fastapi.responses import JSONResponse
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json

 
app = FastAPI()
 
# Load Google Sheets credentials
credentials = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
 
# The credentials file should contain the appropriate service account key JSON
credentials_path = "credentials.json"
 
try:
    credentials = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=SCOPES
    )
 
    if credentials.expired:
        credentials.refresh(Request())
 
except exceptions.GoogleAuthError as e:
    raise HTTPException(status_code=500, detail=str(e))
 
gc = gspread.authorize(credentials)
 
# Specify your Google Sheet name
sheet_name = "SPC"

SMTP_USER = "developer@agilecyber.com"
SMTP_PASSWORD = "sleiemskccrmtmdc"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, to_email, msg.as_string())

def find_empty_row(sheet):
    values = sheet.get_all_values()
    # Find the first empty row
    for i, row in enumerate(values, start=1):
        if not any(row):
            return i
    return len(values) + 1
 
@app.post("/update_sheet/{column_index}")
def update_sheet(
    column_index: int,
    value: dict = Body(...),
    email: str = Body(...),
):
    try:
        valueV = value['data']

        if not valueV or not email:
            raise HTTPException(status_code=400, detail="Invalid payload format")

        # Open the Google Sheet
        spreadsheet = gc.open(sheet_name)
        sheet = spreadsheet.worksheet("spc-questions")

        row = find_empty_row(sheet)

        responses = []

        for item in valueV:
            question_key = list(item.keys())[0]
            answers = item[question_key]

            # Extract question and answers from the provided JSON data
            row = find_empty_row(sheet)
            sheet.update_cell(row, column_index, f"Question: {question_key}\nAnswers: {', '.join(answers)}")

            # Add response to the list
            responses.append(f"Question: {question_key}\nAnswers: {', '.join(answers)}")

        # Join responses into a single string
        report = "\n\n".join(responses)

        # Send email notification with the report
        subject = "Self Performance Checker Notification"
        body = f"Responses:\n\n{report}"
        send_email(email, subject, body)


        return {"message": f"Your response for {question_key} has been recorded and sent to {email}"}
 
    except gspread.exceptions.APIError as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/view_sheet")
def view_sheet(lang: str = 'en'):
    try:
        # Open the Google Sheet
        spreadsheet = gc.open(sheet_name)

        # Determine the worksheet based on language
        worksheet_name = "questions-japanese" if lang.lower() == "ja" else "questions-english"

        # Get the desired sheet
        sheet = spreadsheet.worksheet(worksheet_name)

        # Get all values from the sheet
        all_values = sheet.get_all_values()

        # Convert the values to a list of dictionaries
        headers = all_values[0]
        data = [dict(zip(headers, row)) for row in all_values[1:]]

        return JSONResponse(content={"headers": headers, "data": data}, status_code=200)

    except gspread.exceptions.APIError as e:
        raise HTTPException(status_code=500, detail=str(e))