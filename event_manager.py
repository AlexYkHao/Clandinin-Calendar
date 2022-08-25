import os.path
import pandas as pd
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class EventManager(object):

    def __init__(self, token):
        self.TIME_ZONE = 'America/Los_Angeles'
        self.SCOPES = ['https://www.googleapis.com/auth/calendar.events']
        self.calendar = 'primary'
        self.creds = None
        self.events_df = None
        self.events = []
        if os.path.exists(token):
            self.creds = Credentials.from_authorized_user_file('creds/token.json', self.SCOPES)
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                print("Token doesn't exsits or has expired!")
        try:
            self.service = build('calendar', 'v3', credentials=self.creds)
        except HttpError as error:
            print('An error occurred: %s' % error)

    def read_excel(self, excel_path):
        excel_data = pd.read_excel(excel_path)
        self.events_df = pd.DataFrame(excel_data,
                                      columns=['Date', 'Start Time', 'End Time', 'Event Title', 'Location',
                                               'Meeting Type'])

    def generate_future_events(self):
        if self.events_df is None:
            return
        for index, row in self.events_df.iterrows():
            body = {"summary": None,
                    "start": None,
                    "end": None,
                    "location": None,
                    "description": None}
            try:
                event_date = row['Date'].to_pydatetime()
                start_time = row['Start Time']
                start_time = datetime.combine(event_date, start_time).isoformat()
                end_time = row['End Time']
                end_time = datetime.combine(event_date, end_time).isoformat()
                body["start"] = {"dateTime": start_time, "timeZone": self.TIME_ZONE}
                body["end"] = {"dateTime": end_time, "timeZone": self.TIME_ZONE}
                body["summary"] = row['Event Title']
                body["location"] = row['Location']
                body["description"] = row['Meeting Type']
            except:
                print("Missing critical info for an event!!!")

            if (body["summary"] is not None) and (body["start"] is not None) and (body["end"] is not None):
                if datetime.fromisoformat(body["start"]["dateTime"]) > datetime.now():
                    if self.event_is_new(body):
                        self.events.append(body)

    def event_is_new(self, body):
        all_events = [event["start"] for event in self.events]
        if "start" in body:
            if not body["start"] in all_events:
                return True
            else:
                return False
        else:
            return False

    def remove_all_future_events(self):
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        events_result = self.service.events().list(calendarId=self.calendar, timeMin=now,
                                                   singleEvents=True,  # maxResults=50,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        if events:
            for event in events:
                self.service.events().delete(
                    calendarId=self.calendar,
                    eventId=event['id'],
                ).execute()

    def setup_future_events(self):
        self.remove_all_future_events()
        self.generate_future_events()
        if len(self.events) > 0:
            for body in self.events:
                event_result = self.service.events().insert(calendarId=self.calendar, body=body).execute()
