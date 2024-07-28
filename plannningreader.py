from datetime import date, timedelta, time, datetime
import pymupdf
import os.path
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from backports.zoneinfo import ZoneInfo
from pathlib import Path
from dateutil.parser import parse
import re
import base64
from glob import MAIL_FILTER, PDFS_PATH, CALENDAR_NAME, NAME


SCOPES = ['https://www.googleapis.com/auth/calendar', "https://www.googleapis.com/auth/gmail.readonly"]

CREDENTIALS_FILE = 'credentials.json'

CALENDAR_ID = None

def get_services():
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

   calendar_service = build('calendar', 'v3', credentials=creds)
   mail_service = build("gmail", "v1", credentials=creds)

   return calendar_service, mail_service

def get_all_cal(service):
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

def create_event(service, location, t, d, calid):

    start = datetime(year = d.year, month = d.month, day = d.day, hour = t[0].hour, minute = t[0].minute, tzinfo = ZoneInfo("America/Vancouver"))
    end = datetime(year = d.year, month = d.month, day = d.day, hour = t[1].hour, minute = t[1].minute, tzinfo = ZoneInfo("America/Vancouver"))

    print(location, d, start, end)

    if CALENDAR_ID:
        event_result = service.events().insert(calendarId=calid,
                                            body={
                                                "summary": location,
                                                "location": location,
                                                "start": {"dateTime": start.isoformat(), "timeZone": 'PST'},
                                                "end": {"dateTime": end.isoformat(), "timeZone": 'PST'},
                                            }
        ).execute()
        print("Uploaded")

        return event_result['id']

def get_reception_year(mail):
    for h in mail["payload"]["headers"]:
        if "Date" in h["name"]:
            return parse(h["value"]).year

def download_new_schedules(service):
    downloaded_files = []
    results = mail_service.users().messages().list(userId="me", q=MAIL_FILTER).execute()
    
    for msg in results["messages"]:
        mail = mail_service.users().messages().get(userId="me", id = msg["id"], format="full").execute()
        year = get_reception_year(mail)
        for part in mail["payload"]["parts"]:
            if not "filename" in part:
                continue

            attachement_filename = part["filename"]
            
            if not attachement_filename or Path(PDFS_PATH, attachement_filename).exists():
                continue

            if not "DEPOT" in attachement_filename:
                # Skip attachement witout DEPOT in the filename
                continue
            
            attachment_id = part["body"]["attachmentId"]

            at = mail_service.users().messages().attachments().get(userId="me", messageId = msg["id"], id=attachment_id).execute()

            bat = base64.urlsafe_b64decode(at['data'].encode('UTF-8'))
            with open(f"{Path(PDFS_PATH, attachement_filename)}", "bw") as f:
                f.write(bat)

            print((attachement_filename, year))
            downloaded_files.append((attachement_filename, year))

    return downloaded_files

def get_shifts(pdffile):
    doc = pymupdf.open(Path(PDFS_PATH, pdffile)) # open a document
    for page in doc: # iterate the document pages
        text = page.get_text().encode("utf8") # get plain text (is in UTF-8)

    page = doc[0] # get the 1st page of the document
    tabs = page.find_tables() # locate and extract any tables on page

    for line in tabs[0].extract():
        line[0] = line[0].replace(" ", "") # Clean the name extract from pdf
        if NAME in line:
            return line[1:]

def abrv2timeandloc(abrv):
    location = ""
    if "N" in abrv:
        location = "Nesters Depot"
    if "F" in abrv:
        location = "Function Junction Depot"
    if "SL" in abrv:
        location = "Squamish Landfill"
        t = (time(hour = 8), time(hour = 16))
    if abrv in ["WTS", "WT"] :
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

def get_start_date_from_filename(filename, year):
    m = re.search("\w* [0-9]+", filename, flags=0)
    if m:
        da = parse(m.group(0), ignoretz = True)
        da = datetime(year = year, month = da.month, day = da.day)
        return da
    else:
        return None


if __name__ == "__main__":

    cals_service, mail_service = get_services()

    cals = get_all_cal(cals_service)

    for summary, calid, primary in cals:
        if summary == CALENDAR_NAME:
            CALENDAR_ID = calid

    if not CALENDAR_ID:
        print(f"Could not find Calendar: {CALENDAR_NAME}")
        exit(1)

    filenames = download_new_schedules(mail_service)

    for filename, year in filenames:
        shifts = get_shifts(filename)
        d = get_start_date_from_filename(filename, year)

        for shift in shifts:
            if shift:
                location, t = abrv2timeandloc(shift)
                print(location, t, d)

                create_event(cals_service, location, t, d, CALENDAR_ID)
            d = d + timedelta(days = 1)