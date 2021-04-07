import sys, os, inspect
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
oneUp = os.path.dirname(currentdir)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(oneUp)
from CalendarAPI.calendar import *
from datetime import *
from config import *
import math

if __name__=="__main__":
    kalender = google_calendar(afvalKalender)
    dates1, shiften1 = kalender.createCycle(name_PMD, startDate_PMD, endDate, startCycle_PMD, period_PMD)
    dates2, shiften2 = kalender.createCycle(name_HUISVUIL, startDate_HUISVUIL, endDate, startCycle_HUISVUIL, period_HUISVUIL)
    dates3, shiften3 = kalender.create_monthCycle(name_Papier, startDate_PAPIER, endDate, startCycle_PAPIER, 28, 7)
    dates1.extend(dates2)
    shiften1.extend(shiften2)
    dates1.extend(dates3)
    shiften1.extend(shiften3)
    kalender.createEvents(dates1, shiften1, "")
    kalender.writeCycleToCalendar()
   