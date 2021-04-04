import sys, os, inspect
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
oneUp = os.path.dirname(currentdir)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(oneUp)
from Lib.functions import *
from Lib.getExcelData import *
from Lib.export import *
from CalendarAPI.calendar import *
from datetime import *
from config import *
import math
def writeAfvalKalender():
    kalender = google_calendar(testCalendar)
    startDate = date(2021,1, 7)
    endDate = date(2022, 12, 31)
    dates = []
    shiftenExcel = []
    date_1 = startDate
    d = math.floor((abs(endDate-startDate).days+1)/28)
    prev_month = ""
    for i in range(d):
        if prev_month != date_1.strftime("%m"):
            dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
            prev_month = date_1.strftime("%m")
        else:
            date_1 += timedelta(days=7)
            dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
        date_1 += timedelta(days=28)
        shiftenExcel.append("Papier & Karton")
    kalender.writeEventsToCalendar(dates, shiftenExcel)
    startDate = date(2021,1, 8)
    endDate = date(2022, 12, 31)
    dates = []
    shiftenExcel = []
    date_1 = startDate
    d = math.floor((abs(endDate-startDate).days+1)/14)
    for i in range(d):
        dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
        date_1 += timedelta(days=14)
        shiftenExcel.append("Huisvuil")
    kalender.writeEventsToCalendar(dates, shiftenExcel)
    startDate = date(2021,1,15)
    endDate = date(2022, 12, 31)
    dates = []
    shiftenExcel = []
    date_1 = startDate
    d = math.floor((abs(endDate-startDate).days+1)/14)
    for i in range(d):
        dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
        date_1 += timedelta(days=14)
        shiftenExcel.append("PMD + GFT")
    kalender.writeEventsToCalendar(dates, shiftenExcel)
writeAfvalKalender()