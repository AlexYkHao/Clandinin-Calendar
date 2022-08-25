import os.path
import pandas as pd
from datetime import datetime, timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.events']
TIME_ZONE = 'America/Los_Angeles'

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('creds/token.json'):
        print('json exists!')
        creds = Credentials.from_authorized_user_file('creds/token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creds/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('creds/token.json', 'w') as token:
            token.write(creds.to_json())

    excel_data = pd.read_excel('test/test.xlsx')
    df = pd.DataFrame(excel_data, columns=['Date', 'Start Time', 'End Time', 'Event Title', 'Location', 'Meeting Type'])

    try:
        service = build('calendar', 'v3', credentials=creds)
        for index, row in df.iterrows():
            body = {}
            event_date = row['Date'].to_pydatetime()
            start_time = row['Start Time']
            start_time = datetime.combine(event_date, start_time).isoformat()
            end_time = row['End Time']
            end_time = datetime.combine(event_date, end_time).isoformat()
            body['start'] = {"dateTime": start_time, "timeZone": TIME_ZONE}
            body['end'] = {"dateTime": end_time, "timeZone": TIME_ZONE}
            body['summary'] = row['Event Title']
            body['location'] = row['Location']
            body['description'] = row['Meeting Type']
            print(body)
            event_result = service.events().insert(calendarId='primary', body=body).execute()

            print("created event")
            print("id: ", event_result['id'])

    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()
