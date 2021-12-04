from __future__ import print_function
import datetime, pandas, pytz
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

class GCalendar:

    def __init__(self, credentials="credentials.json"):
        self.credentials = credentials
        self.service = self.generate_service()

    def generate_service(self):
        """
        Gets Token. Returns service.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        return build('calendar', 'v3', credentials=creds)


    def print_future_events(self, max_results=10):
        if not self.service:
            self.service = self.generate_service()
        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('Getting the upcoming 10 events')
        events_result = self.service.events().list(calendarId='primary', timeMin=now,
                                            maxResults=max_results, singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])

    def insert_event(self, event=None):
        if not self.service:
            self.service = self.generate_service()
        if not event:
            raise Exception("Please insert event")

        event = self.service.events().insert(calendarId='primary', body=event).execute()
        print( 'Event created: %s' % (event.get('htmlLink')))


if __name__ == '__main__':
    cal = GCalendar()

    xls = pandas.ExcelFile('southshore.xlsx')
    df = xls.parse(xls.sheet_names[0])
    df.loc[df['Vikings Team'] == 'Girls', 'Start Time'] = df['Date'] + datetime.timedelta(hours=17) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Boys', 'Start Time'] = df['Date'] + datetime.timedelta(hours=17) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Girls Jr.', 'Start Time'] = df['Date'] + datetime.timedelta(hours=15) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Boys Jr.', 'Start Time'] = df['Date'] + datetime.timedelta(hours=15) # Varsity always starts at 5

    df.loc[df['Vikings Team'] == 'Girls', 'End Time'] = df['Date'] + datetime.timedelta(hours=18) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Boys', 'End Time'] = df['Date'] + datetime.timedelta(hours=18) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Girls Jr.', 'End Time'] = df['Date'] + datetime.timedelta(hours=16) # Varsity always starts at 5
    df.loc[df['Vikings Team'] == 'Boys Jr.', 'End Time'] = df['Date'] + datetime.timedelta(hours=16) # Varsity always starts at 5
    for index, row in df.iterrows():
        event = {
                'summary': f'[{"Varsity" if not row["Vikings Team"].endswith(".") else "Junior"}] {row["Vikings Team"]} vs {row["Other Team"]}',
                'location': f"{row['Location']}",
                'description': 'Vikings contract to cover basketball games',
                'start': {
                    'dateTime': row['Start Time'].isoformat(),
                    'timeZone': 'America/New_York',
                },
                'end': {
                    'dateTime': row['End Time'].isoformat(),
                    'timeZone': 'America/New_York',
                },
                'recurrence': [
                ],
                'attendees': [
                ],
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': 60},
                    {'method': 'popup', 'minutes': 30},
                    ],
                },
            }
        cal.insert_event(event)
    cal.print_future_events(10000)

