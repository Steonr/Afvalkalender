import sys, os, inspect
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
oneUp = os.path.dirname(currentdir)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(oneUp)
from CalendarAPI.calendar import *
from datetime import *
from config import *
import math
'''d = 1
dates = []
shiftenExcel = []
prev_month = ""
date_1 = date(2021,1, 7)
for i in range(d):
        if prev_month != date_1.strftime("%m"):
            dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
            prev_month = date_1.strftime("%m")
        else:
            date_1 += timedelta(days=7)
            dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
        date_1 += timedelta(days=28)
        shiftenExcel.append("Papier & Karton")
kalender.createEvents(dates, shiftenExcel)'''

kalender = google_calendar(testCalendar)
dates1, shiften1 = kalender.createCycle(name_PMD, startDate_PMD, endDate, startCycle_PMD, period_PMD)
dates2, shiften2 = kalender.createCycle(name_HUISVUIL, startDate_HUISVUIL, endDate, startCycle_HUISVUIL, period_HUISVUIL)

dates1.extend(dates2)
shiften1.extend(shiften2)
kalender.createEvents(dates1, shiften1, "")
kalender.writeCycleToCalendar()

if __name__=="__main__":
    pass