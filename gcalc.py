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
import httplib2
import socks

import argparse
import httplib2
import os
import time
from pprint import pprint

from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import hashlib

from time import gmtime, strftime
from datetime import date, time, datetime, timedelta

from googleapiclient import sample_tools

def init(argv, name, version, doc, filename, scope=None, parents=[], http_proxy=None):
  """A common initialization routine for samples.

  Many of the sample applications do the same initialization, which has now
  been consolidated into this function. This function uses common idioms found
  in almost all the samples, i.e. for an API with name 'apiname', the
  credentials are stored in a file named apiname.dat, and the
  client_secrets.json file is stored in the same directory as the application
  main file.

  Args:
    argv: list of string, the command-line parameters of the application.
    name: string, name of the API.
    version: string, version of the API.
    doc: string, description of the application. Usually set to __doc__.
    file: string, filename of the application. Usually set to __file__.
    parents: list of argparse.ArgumentParser, additional command-line flags.
    scope: string, The OAuth scope used.

  Returns:
    A tuple of (service, flags), where service is the service object and flags
    is the parsed command-line flags.
  """
  if scope is None:
    scope = 'https://www.googleapis.com/auth/' + name

  # Parser command-line arguments.
  parent_parsers = [tools.argparser]
  parent_parsers.extend(parents)
  parser = argparse.ArgumentParser(
      description=doc,
      formatter_class=argparse.RawDescriptionHelpFormatter,
      parents=parent_parsers)
  flags = parser.parse_args(argv[1:])

  # Name of a file containing the OAuth 2.0 information for this
  # application, including client_id and client_secret, which are found
  # on the API Access tab on the Google APIs
  # Console <http://code.google.com/apis/console>.
  client_secrets = os.path.join(os.path.dirname(filename),
                                'client_secrets.json')

  # Set up a Flow object to be used if we need to authenticate.
  flow = client.flow_from_clientsecrets(client_secrets,
      scope=scope,
      message=tools.message_if_missing(client_secrets))

  # Prepare credentials, and authorize HTTP object with them.
  # If the credentials don't exist or are invalid run through the native client
  # flow. The Storage object will ensure that if successful the good
  # credentials will get written back to a file.
  storage = file.Storage(name + '.dat')
  credentials = storage.get()
  if credentials is None or credentials.invalid:
    credentials = tools.run_flow(flow, storage, flags)
 
  http = credentials.authorize(http = httplib2.Http(proxy_info = http_proxy))

  # Construct a service object via the discovery service.
  service = discovery.build(name, version, http=http)
  return (service, flags)
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

def quickAddEvent(service, calendarID, summary):
    # Quick Add
    created_event = service.events().quickAdd(
            calendarId=calendarID,
            text=summary).execute()

    print created_event['id']

def genID(summary, startTime, stopTime, location):
	return hashlib.md5(summary + str(startTime) + str(stopTime) + location).hexdigest()
	
def addEvent(service, calendarID, summary, location=None, 
							startTime=None, stopTime=None):
	event = {
			'id': genID(summary, startTime, stopTime, location),
            'summary': summary,
            'location': location,
			'start': {
                'dateTime' : formatTime(startTime) 
            },
            'end': {
                'dateTime' : formatTime(stopTime) 
            },
        }
	try:
		created_event = service.events().insert(calendarId=calendarID, body=event).execute()
		print "Adding new Event : " + event['summary']
	except errors.HttpError, e:
		err = json.loads(e.content)
		err_msg = err['error']['message']
		# Print err message other than duplicates
		if not "already exists" in err_msg:
			print err_msg
	
def formatTime(date=None):
		return date.strftime("%Y-%m-%dT%H:%M:%S+02:00")
	
def now(gmt=0):
    d = datetime.now() + timedelta(hours=gmt)
    return d.strftime("%Y-%m-%dT%H:%M:%S+02:00")

def main(argv):

	# TODO
	#http_proxy=None
	http_proxy=httplib2.ProxyInfo(socks.PROXY_TYPE_HTTP, 'proxy.jf.intel.com', 911)
	#print http_proxy

    # Authenticate and construct service.
    #service, flags = sample_tools.init(
	service, flags = init(
        argv, 'calendar', 'v3', __doc__, __file__,
        scope='https://www.googleapis.com/auth/calendar', http_proxy=http_proxy)

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
