#!/usr/bin/env python3

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Delete the file token.pickle whenever these scopes are modified
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of the spreadsheet to upload to

# Test spreadsheet
#SPREADSHEET_ID = "1JB6cq7TQAs0TbYPqCifX1VmyAtG3WYorRzjWLSBmyNs"
#RANGE = "Sheet1"

# Filament Testing Specimen sheet
SPREADSHEET_ID = "1tVri6CrL-bwJ5yE0rWfuB90tk2gGPgzlQS7GmYKdL9o"
RANGE = "ASTM D790 Test Data - Raw"

def login():
    """Checks for existing credentials and logs into Google if required.
    """

    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorizations flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open ('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no valid credentials available, let the user log in.
    if not creds or not creds.valid:
        print(':: Login Required')
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def upload(creds, data):
    """Appends the supplied data to the Google Sheet.
    """

    print(':: Uploading Tensile Data to Google')

    service = build('sheets', 'v4', credentials=creds)
    body = { 'values': data }
    result = service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range=RANGE,
        valueInputOption='USER_ENTERED', body=body
    ).execute()


def main():
    # Append data to sheet
    values = [
        [1,2,3],
        [4,5,6]
    ]

    creds = login()
    upload(creds, values)


if __name__ == '__main__':
    main()
