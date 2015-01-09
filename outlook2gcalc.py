"""
    Print Outlook events for the next month starting from today, 
    for the default Outlookd calendar.

    Dependencies: win32com package 
"""


import win32com.client
import pywintypes

import time
import datetime

import gcalc
import socks
import httplib2
import sys

from pprint import pprint

def convertTime(pytime):
	return datetime.datetime.fromtimestamp (int (pytime))

def main(argv):
	## INIT Outlook
	outlook = win32com.client.Dispatch("Outlook.Application")
	namespace = outlook.GetNamespace("MAPI")
	appointments = namespace.GetDefaultFolder(9).Items

	appointments.Sort("[Start]")
	appointments.IncludeRecurrences = "True"

	# Restrict to items in the next day
	begin = datetime.date.today()
	end = begin + datetime.timedelta(days = 30);
	restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
	restrictedItems = appointments.Restrict(restriction)

	## INIT Gcalc
	http_proxy=httplib2.ProxyInfo(socks.PROXY_TYPE_HTTP, 'proxy.jf.intel.com', 911)
	#print http_proxy

	# Authenticate and construct service.
	#service, flags = sample_tools.init(
	service, flags = gcalc.init(
		argv, 'calendar', 'v3', __doc__, __file__,
		scope='https://www.googleapis.com/auth/calendar', http_proxy=http_proxy)

	# Get calendar id by name (summary)
	calendarID = gcalc.getCalendarID(service, 'WORK')
	#calendarID = getCalendarID(service, 'FUN STUFF')
	if calendarID == None:
		print ('No calendar was found')
		sys.exit() 
	calendar =  service.calendars().get(calendarId=calendarID).execute()


	# Iterate through restricted AppointmentItems and print them
	for appointmentItem in restrictedItems:
			summary = appointmentItem.Subject.encode('ascii', 'ignore')
			organizer = appointmentItem.Organizer.encode('ascii', 'ignore')
			location = appointmentItem.Location.encode('ascii', 'ignore')
			
			start = convertTime(appointmentItem.Start)
			end = convertTime(appointmentItem.End)
			
			gcalc.addEvent(service, calendarID, summary, location, start, end)

if __name__ == '__main__':
    main(sys.argv)