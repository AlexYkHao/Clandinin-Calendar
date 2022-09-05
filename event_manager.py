import hashlib
import os.path
import pandas as pd
from datetime import datetime
from pytz import timezone

from openpyxl import load_workbook
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

"""
Goal:
Add, update and delete lab meeting events on calendar based on excel sheets.
Detailed demand:
1, add events: when there is another events that's not on the excel sheets, remove it and add the requested one
2, update events: update the info of an existing calendar
3, delete events: remove events that are no longer in the excel sheets but are still on the calendar
Logistics:
1, Pull all the future events from calendar, find the ones that are created by the manager (cal_events)
2, Read the excel files in sequence to create an event list (excel_events)
3, Find the events to be removed, and remove them
4, Find the events to be updated, and update them, update the excel
5, Find the events to be created, and create them, update the excel
6, If anything went wrong, can delete all the future events created by a certain user/managerID
"""


class ExcelHandler(object):

    def __init__(self, excel_path, col_mapper):
        self.excel_path = excel_path
        if (not os.path.exists(self.excel_path)) and ('.xlsx' not in self.excel_path):
            print('Wrong excel path!!!')
            return
        self.wb = load_workbook(filename=excel_path)
        self.sheet = self.wb.active
        self.col_mapper = col_mapper

    def update_cell(self, row, col_head, var):
        if col_head in self.col_mapper.keys():
            row = row + 2
            col = self.col_mapper[col_head]
            cell_ = col + str(row)
            self.sheet[cell_] = var

    def close_book(self):
        self.wb.save(filename=self.excel_path)


class EventManager(object):

    def __init__(self, token, manager='manager_meet', excel_path=''):
        self.TIME_ZONE = 'America/Los_Angeles'
        self.SCOPES = ['https://www.googleapis.com/auth/calendar.events']
        self.calendar = 'primary'
        self.manager_identifier = manager  # this is a string for identifying events that's created by EventManager
        self.creds = None
        self.events_df = None
        self.excel_writer = None
        self.excel_path = excel_path
        self.excel_events = []
        self.events_to_add = []
        self.events_to_delete = []
        self.events_to_update = []
        self.cal_events = []
        if (not os.path.exists(self.excel_path)) and ('.xlsx' not in self.excel_path):
            print('Wrong excel path!!!')
            return
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

    def pull_calendar_events(self):
        self.cal_events = []
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        events_result = self.service.events().list(calendarId=self.calendar, timeMin=now,
                                                   singleEvents=True,  # maxResults=50,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        if events:
            for event in events:
                if self.manager_identifier in event['id']:
                    self.cal_events.append(event)

    def read_excel(self, col_mapper):
        excel_data = pd.read_excel(self.excel_path)
        self.events_df = pd.DataFrame(excel_data,
                                      columns=['ID', 'Date', 'Start Time', 'End Time', 'Event Title', 'Location',
                                               'Meeting Type', 'Error State', 'Need Update'])
        self.add_excel_events()
        self.excel_writer = ExcelHandler(excel_path=self.excel_path, col_mapper=col_mapper)

    def add_excel_events(self):
        self.excel_events = []  # so far I have to keep this since 1 excel can be processed at a time
        if self.events_df is None:
            return
        for index, row in self.events_df.iterrows():
            body = {"id": None,
                    "summary": None,
                    "start": None,
                    "end": None,
                    "location": None}
            try:
                event_date = row['Date'].to_pydatetime()
                start_time = row['Start Time']
                start_time = datetime.combine(event_date, start_time).isoformat()
                end_time = row['End Time']
                end_time = datetime.combine(event_date, end_time).isoformat()
                body["id"] = row['ID']
                body["start"] = {"dateTime": start_time, "timeZone": self.TIME_ZONE}
                body["end"] = {"dateTime": end_time, "timeZone": self.TIME_ZONE}
                body["summary"] = row['Event Title']
                body["location"] = row['Location']
                # body["description"] = row['Meeting Type']
            except:
                print("Missing critical field in excel for an event!!!")

            if (body["summary"] is not None) and (body["start"] is not None) and (body["end"] is not None):
                if datetime.fromisoformat(body["start"]["dateTime"]) > datetime.now():
                    if self.event_is_new(body):
                        event = {
                            'body': body,
                            'id': body['id'],
                            'update': row['Need Update'],
                            'row_index': index
                        }
                        self.excel_events.append(event)

    def event_is_new(self, body):
        all_events = [event['body']["start"] for event in self.excel_events]
        if "start" in body:
            if not body["start"] in all_events:
                return True
            else:
                return False
        else:
            return False

    def generate_new_events(self):
        self.events_to_add = []
        for event in self.excel_events:
            if pd.isna(event['id']):
                self.events_to_add.append(event)

    def find_events_to_update(self):
        self.events_to_update = []
        for event in self.excel_events:
            if not (pd.isna(event['id']) and pd.isna(event['update'])):
                self.events_to_update.append(event)

    def find_events_to_delete(self):
        self.events_to_delete = []
        self.pull_calendar_events()
        excel_ids = [event['id'] for event in self.excel_events if not pd.isna(event['id'])]
        for event in self.cal_events:
            if not (event['id'] in excel_ids):
                self.events_to_delete.append(event)

    def delete_events(self):
        self.find_events_to_delete()
        for event in self.events_to_delete:
            try:
                self.service.events().delete(
                    calendarId=self.calendar,
                    eventId=event['id'],
                ).execute()
            except HttpError:
                print('connection error while deleting!')

    def update_events(self):
        self.find_events_to_update()
        for event in self.events_to_update:
            try:
                event_result = self.service.events().update(
                    calendarId=self.calendar,
                    eventId=event['id'],
                    body=event['body']
                ).execute()
                # self.events_df.loc[event['row_index'], 'Need Update'] = None
                self.excel_writer.update_cell(row=event['row_index'], col_head='Need Update', var=None)
            except HttpError:
                print('connection error while updating!')

    def remove_overlap(self, event):
        try:
            start_time = datetime.fromisoformat(event["start"]["dateTime"])
            start_time = timezone('America/Los_Angeles').localize(start_time).isoformat()
            end_time = datetime.fromisoformat(event["end"]["dateTime"])
            end_time = timezone('America/Los_Angeles').localize(end_time).isoformat()
            events_overlap_results = self.service.events().list(calendarId=self.calendar,
                                                                timeMin=start_time,
                                                                timeMax=end_time,
                                                                singleEvents=True,
                                                                orderBy='startTime').execute()
            events_overlap = events_overlap_results.get('items', [])
            for event in events_overlap:
                self.service.events().delete(
                    calendarId=self.calendar,
                    eventId=event['id'],
                ).execute()
        except HttpError:
            print("connection error while removing overlaps!")

    def generate_id(self, body):
        id_string = body["summary"] + body["start"]["dateTime"] + "created" + datetime.now().isoformat()
        id_ = self.manager_identifier + hashlib.sha224(id_string.encode()).hexdigest()
        return id_

    def add_events(self, allow_overlap=False):
        self.generate_new_events()
        for event in self.events_to_add:
            try:
                if not allow_overlap:
                    self.remove_overlap(event['body'])
                body = event['body']
                body['id'] = self.generate_id(body)
                event_result = self.service.events().insert(calendarId=self.calendar, body=body).execute()
                # self.events_df.loc[event['row_index'], 'ID'] = event_result['id']
                self.excel_writer.update_cell(row=event['row_index'], col_head='ID', var=event_result['id'])
            except HttpError:
                print('connection error while adding events!')
                # self.events_df.loc[event['row_index'], 'Error State'] = 1
                self.excel_writer.update_cell(row=event['row_index'], col_head='Error State', var=1)

    def update_from_excel(self, col_mapper, allow_overlap=False):
        self.read_excel(col_mapper)
        self.delete_events()
        self.update_events()
        self.add_events(allow_overlap)
        self.excel_writer.close_book()

    def remove_events_created_by(self, email):
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        events_result = self.service.events().list(calendarId=self.calendar, timeMin=now,
                                                   singleEvents=True,  # maxResults=50,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        if events:
            for event in events:
                if event["creator"]["email"] == email:
                    self.service.events().delete(
                        calendarId=self.calendar,
                        eventId=event['id'],
                    ).execute()

    def remove_events_id_with(self, id_prefix):
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        events_result = self.service.events().list(calendarId=self.calendar, timeMin=now,
                                                   singleEvents=True,  # maxResults=50,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        if events:
            for event in events:
                if id_prefix in event["id"]:
                    self.service.events().delete(
                        calendarId=self.calendar,
                        eventId=event['id'],
                    ).execute()
