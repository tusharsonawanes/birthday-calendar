from __future__ import print_function

import os.path
import pandas as pd
import datetime
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Reading excel
dataframe = pd.read_excel('data/google_data.xlsx', 'new', skiprows=0)   #enter the excel file location and name, followed by the sheet name

def main():
    count=0
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
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

    try:
        service = build('calendar', 'v3', credentials=creds)

        for i in range(len(dataframe)): 
            member_dob_temp = pd.to_datetime(dataframe["Member DOB - "].values[count])
            member_dob = datetime.date(member_dob_temp)
            member_name = "{}".format(dataframe["Member Name - "].values[count])
            member_instagram = "{}".format(dataframe["Member Instagram ID - "].values[count])
            print(member_dob)
            print(member_name)
            print(member_instagram)
            event = {
                'summary': "{}'s Birthday".format(member_name),
                'location': 'Instagram',
                'description': 'Please upload a picture for {}, \nInstagram ID - {}'.format(member_name, member_instagram),
                'start': {
                  'date': '{}'.format(member_dob),
                  'timeZone': 'Asia/Kolkata',     # Can be replaced with your timezone
                },
                'end': {
                  'date': '{}'.format(member_dob),
                  'timeZone': 'Asia/Kolkata',     # Can be replaced with your timezone
                },
                'recurrence': [
                  'RRULE:FREQ=YEARLY'
                ],
                'attendees': [
                  {'email': 'mail@expaple.com'},  # Email ID of the first recepient to receive notification
                  {'email': 'mail@example@.com'}  # Email ID of the second recepient to receive notification
                ],
                'reminders': {
                  'useDefault': False,
                  'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 10},
                  ],
                },
            }
            event = service.events().insert(calendarId='primary', body=event).execute()
            print('Event created: %s'%(event.get('htmlLink')))
            count=count+1            
            time.sleep(2)

    except HttpError as error:
        print('An error occurred: %s' % error)

if __name__ == '__main__':
    main()