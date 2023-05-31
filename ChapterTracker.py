import os
import spreadsheetID
import pandas as pd
import tkinter as tk
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = spreadsheetID.testsheet

class GUI():
    
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title('ChapterTracker')

        self.memberNames = open('memberNames.csv')
        self.checkInList = None
        self.excusedList = None

        
        self.root.mainloop()

def record(service):
    checkIn = pd.read_csv('')
    names = pd.read_csv('')
    excused = pd.read_csv('')

    n = names['Names'].to_list()
    ch = checkIn['Full Name'].to_list()
    exName = excused['Full Name'].to_list()

    for i in range(len(n)):
        n[i] = n[i].lower().strip()

    for i in range(len(ch)):
        ch[i] = ch[i].lower().strip()

    for i in range(len(exName)):
        exName[i] = exName[i].lower().strip()

    missing = []
    absent = []

    for i in n:
        if (i not in ch) and (i not in exName):
            absent.append(i)
            print(i)

    for i in ch:
        if i not in n:
            missing.append(i)

    print("Missing names:", missing)
    print("Excused:", exName)

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        record(sheets)

        for row in range(2, 8):
            num1 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!A{row}").execute().get("values")[0][0])
            num2 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!B{row}").execute().get("values")[0][0])
            calc_result = num1 + num2
            print(f"Processing {num1} + {num2}")

            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!C{row}",
                valueInputOption="USER_ENTERED", body={"values": [[f"{calc_result}"]]}).execute()

            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!D{row}",
                valueInputOption="USER_ENTERED", body={"values": [[f"Done"]]}).execute()

    except HttpError as error:
        print(error)


if __name__ == "__main__":
    GUI()
