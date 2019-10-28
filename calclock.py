from __future__ import print_function
import datetime
import dateutil.parser
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

def is_ascii(s):
    return all(ord(c) < 128 for c in s)

def isValidAlarm(event, minutesBeforeEvent, workingHours):
    valid = True
    if not is_ascii(event['summary']):
        valid = False
    start = event['start'].get('dateTime', event['start'].get('date'))
    start = dateutil.parser.parse(start) - datetime.timedelta(minutes=minutesBeforeEvent)
    if start.hour < workingHours['start'] or start.hour > workingHours['end']:
        valid = False
    return valid

def getService(scopes, tokenFileName, credFileName):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(tokenFileName):
        with open(tokenFileName, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credFileName, scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(tokenFileName, 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)

def getEvents(service, maxResults=10):
    now = datetime.datetime.now() - datetime.timedelta(hours=2) # 'Z' indicates UTC time
    events_result = service.events().list(calendarId='primary', timeMin=now.isoformat()+ 'Z',
                                        maxResults=maxResults, singleEvents=True,
                                        orderBy='startTime').execute()
    return events_result.get('items', [])

def create_event(service, start_time_str, summary, duration=1,attendees=None, description=None, location=None):
    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': start_time_str[:-7],
            'timeZone': 'Asia/Jerusalem',
        },
        'end': {
            'dateTime': start_time_str[:-7],
            'timeZone': 'Asia/Jerusalem',
        }
    }

    return service.events().insert(calendarId='primary', body=event,sendNotifications=False).execute()

def getAlarmEvents(events):
    existingAlarms = []

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        if 'Alarm' in event['summary']:
            existingAlarms.append({'start': start, 'title': event['summary']})

    return existingAlarms

def writeAlarmsEvents(calCfg, alarms):
    alarmsAdded = 0
    service = getService(calCfg['scopes'], calCfg['tokenFile'], calCfg['credFile'])
    existingEvents = getEvents(service)
    existingAlarms = getAlarmEvents(existingEvents)
    
    for alarm in alarms:
        alarmFound = False
        for existedAlarm in existingAlarms:
            if existedAlarm['start'] == alarm['start']:
                alarmFound = True
        if not alarmFound:
            print('Alarm set successfully')
            create_event(service, alarm['start'], 'Alarm')
            alarmsAdded = alarmsAdded+1
    return alarmsAdded
    

# TODO: make alarms as set
def extractAlarmsFrom(events, minutesBeforeEvent, workingHours):
    alarms = []
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        start = dateutil.parser.parse(start) - datetime.timedelta(minutes=minutesBeforeEvent)
        if isValidAlarm(event, minutesBeforeEvent, workingHours):
            alarms.append({'start': start.isoformat(), 'title': 'Alarm'})
    return alarms

def generateAlarms(calCfg, workingHours, minutesBeforeEvent=10):
    service = getService(calCfg['scopes'], calCfg['tokenFile'], calCfg['credFile'])
    events = getEvents(service)
    alarms = extractAlarmsFrom(events, minutesBeforeEvent, workingHours)
    return alarms

def main():
    # calendar 1 and 2 can be the calendar, but I wanted isolation to I used 2.

    calendar1 = {     'scopes': ['https://www.googleapis.com/auth/calendar.readonly']
                    , 'tokenFile': 'token_cal1.pickle'
                    , 'credFile': 'credentials_cal1.json'
                }
    
    calendar2 = {     'scopes': ['https://www.googleapis.com/auth/calendar.events']
                    , 'tokenFile': 'token_cal2.pickle'
                    , 'credFile': 'credentials_cal2.json'
                }

    workingHours = {'start': 11, 'end': 20}
    alarms = generateAlarms(calendar1, workingHours)
    alarmsAdded = writeAlarmsEvents(calendar2, alarms)
    print('{} alarms added'.format(alarmsAdded))

if __name__ == '__main__':
    main()