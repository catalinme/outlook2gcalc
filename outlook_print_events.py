"""
    Print Outlook events for the next month starting from today, 
    for the default Outlookd calendar.

    Dependencies: win32com package 
"""


import win32com.client
import time
import datetime

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
appointments = namespace.GetDefaultFolder(9).Items

appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

# Restrict to items in the next 30 days (using Python 3.3 - might be slightly different for 2.7)
begin = datetime.date.today()
end = begin + datetime.timedelta(days = 30);
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
restrictedItems = appointments.Restrict(restriction)

# Iterate through restricted AppointmentItems and print them
for appointmentItem in restrictedItems:
        print("======= {0}\n Start: {1}, End: {2}, Organizer: {3}\n".format(
                      appointmentItem.Subject, appointmentItem.Start, 
                                appointmentItem.End, appointmentItem.Organizer))
