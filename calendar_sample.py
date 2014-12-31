#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Copyright 2014 Google Inc. All Rights Reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Simple command-line sample for the Calendar API.
Command-line application that retrieves the list of the user's calendars."""

import sys
import json
from time import gmtime, strftime
from datetime import date, time, datetime, timedelta

from oauth2client import client
from googleapiclient import sample_tools

def getCalendarID(service, summary):
    try:
        page_token = None
        while True:
            calendar_list = service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                if calendar_list_entry['summary'] == summary:
                    return calendar_list_entry['id']
                #print calendar_list_entry['summary'] + ' ' + calendar_list_entry['id']
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break
    except client.AccessTokenRefreshError:
        print ('The credentials have been revoked or expired, please re-run'
                'the application to re-authorize.')

def listEvents(service, calendarID):
    #print calendar
    page_token = None
    while True:
        events = service.events().list(calendarId=calendarID, pageToken=page_token, \
                                    #maxResults=100, timeMin='', timeMax='' \
                                    singleEvents='True',orderBy='startTime').execute()
        for event in events['items']:
            try:
                print event['summary']
                print json.dumps(event['start'], indent=4, separators=(',',':'))
            except KeyError:
                print ('No event summary')
                print event
        page_token = events.get('nextPageToken')
        if not page_token:
            break 

def quickAddEvent(service, calendarID):
    # Quick Add
    created_event = service.events().quickAdd(
            calendarId=calendarID,
            text='Test').execute()

    print created_event['id']

def addEvent(service, calendarID):
    event = {
            'summary': 'App',
            'location': 'Somewhere',
            'start': {
                'dateTime': now() 
            },
            'end': {
                'dateTime': now() 
            },
        }
    print event
    created_event = service.events().insert(calendarId=calendarID, body=event).execute()
    print created_event['id'] 

def now(gmt=0):
    d = datetime.now() + timedelta(hours=gmt)
    return d.strftime("%Y-%m-%dT%H:%M:%S+02:00")

def main(argv):
    # Authenticate and construct service.
    service, flags = sample_tools.init(
        argv, 'calendar', 'v3', __doc__, __file__,
        scope='https://www.googleapis.com/auth/calendar')

    # Get calendar id by name (summary)
    calendarID = getCalendarID(service, 'WORK')
    #calendarID = getCalendarID(service, 'FUN STUFF')
    if calendarID == None:
        print ('No calendar was found')
        return 
    calendar =  service.calendars().get(calendarId=calendarID).execute()

    listEvents(service, calendarID)
    addEvent(service, calendarID)

if __name__ == '__main__':
    main(sys.argv)
