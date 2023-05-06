import gspread
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


def write_to_excel(data):
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    gc = gspread.authorize(creds)

    try:
        sh = gc.open('transfer_bank')
    except gspread.SpreadsheetNotFound:
        sh = gc.create('transfer_bank')
    worksheet = sh.get_worksheet(0)
    if len(worksheet.findall('bank')) == 0:
        worksheet.append_row(list(data[0].keys()))
    for row in data:
        drive_number = worksheet.findall(row['number'])
        if len(drive_number) == 0:
            worksheet.append_row(list(row.values()))
