from datetime import date, timedelta, time, datetime
from pypdf import PdfReader
import os.path
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import argparse
from backports.zoneinfo import ZoneInfo

SCOPES = ['https://www.googleapis.com/auth/calendar']

CREDENTIALS_FILE = 'credentials.json'

CALENDAR_ID = None

def get_calendar_service():
   creds = None
   # The file token.pickle stores the user's access and refresh tokens, and is
   # created automatically when the authorization flow completes for the first
   # time.
   if os.path.exists('token.pickle'):
       with open('token.pickle', 'rb') as token:
           creds = pickle.load(token)
   # If there are no (valid) credentials available, let the user log in.
   if not creds or not creds.valid:
       if creds and creds.expired and creds.refresh_token:
           creds.refresh(Request())
       else:
           flow = InstalledAppFlow.from_client_secrets_file(
               CREDENTIALS_FILE, SCOPES)
           creds = flow.run_local_server(port=50000)

       # Save the credentials for the next run
       with open('token.pickle', 'wb') as token:
           pickle.dump(creds, token)

   service = build('calendar', 'v3', credentials=creds)
   return service

def get_all_cal():
    service = get_calendar_service()
    cals = []
    # Call the Calendar API
    print('Getting list of calendars')
    calendars_result = service.calendarList().list().execute()

    calendars = calendars_result.get('items', [])

    if not calendars:
       print('No calendars found.')
    for calendar in calendars:
        summary = calendar['summary']
        id = calendar['id']
        primary = "Primary" if calendar.get('primary') else ""
        cals.append((summary, id, primary))
    return cals

def create_event(location, t, d, calid):
    service = get_calendar_service()

    start = datetime(year = d.year, month = d.month, day = d.day, hour = t[0].hour, minute = t[0].minute, tzinfo = ZoneInfo("America/Vancouver"))
    end = datetime(year = d.year, month = d.month, day = d.day, hour = t[1].hour, minute = t[1].minute, tzinfo = ZoneInfo("America/Vancouver"))

    print(location, d, start, end)

    if args.upload and CALENDAR_ID:
        event_result = service.events().insert(calendarId=calid,
                                            body={
                                                "summary": location,
                                                "location": location,
                                                "start": {"dateTime": start.isoformat(), "timeZone": 'PST'},
                                                "end": {"dateTime": end.isoformat(), "timeZone": 'PST'},
                                            }
        ).execute()

        return event_result['id']

parser = argparse.ArgumentParser(
                    prog='gflplanningreader',
                    description='Reads planning from GFL and uploads it to my google calendar')

parser.add_argument('filepath', help = "Path to PDF planning file to read")
parser.add_argument('startdate', help = "Date of the first day on the schedule")
parser.add_argument('calendar_name', help = "Name of the google calendar to upload the events to.")
parser.add_argument('-u', '--upload', action="store_true", help = "Only upload the events to google calendar if option present.")

args = parser.parse_args()

start_date = date.fromisoformat(args.startdate)

cals = get_all_cal()

for summary, calid, primary in cals:
    if summary == args.calendar_name:
        CALENDAR_ID = calid

reader = PdfReader(args.filepath)
page = reader.pages[0]
for line in page.extract_text().split('\n'):
    if "Romain" in line or "Rom ain" in line:
        shifts = line.replace(" ", ".").split(".")[2:-1]
        while len(shifts) > 14:
            print("To much days, select unnecessary columnn to be deleted:\n", [x for x in zip(range(0, len(shifts)), shifts)])
            text = input("Column index: ")
            if text.isdigit() and int(text) > 0 and int(text) < len(shifts):
                del shifts[int(text)]
            else:
                exit(1)
        print(shifts)

def abrv2timeandloc(abrv):
    location = ""
    if "N" in abrv:
        location = "Nesters Depot"
    if "F" in abrv:
        location = "Function Junction Depot"
    if "SL" in abrv:
        location = "Squamish Landfill"
    if "WTS" in abrv:
        location = "Whistler Transfert Station"
        t = (time(hour = 8), time(hour = 17))
    if "7" in abrv:
        t = (time(hour = 7), time(hour = 13))
    if "1" in abrv:
        t = (time(hour = 13), time(hour = 19))
    if "8" in abrv:
        t = (time(hour = 8), time(hour = 16))
    if "10" in abrv:
        t = (time(hour = 10), time(hour = 16))

    return (location, t)

d = start_date
for shift in shifts:
    if shift:
        location, t = abrv2timeandloc(shift)
        create_event(location, t, d, CALENDAR_ID)
    d = d + timedelta(days = 1)