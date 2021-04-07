import sys, os, inspect
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(THIS_FOLDER)
import math
import pickle
import datetime
from config import *
from datetime import *
from src.Lib.functions import *
from apiclient.discovery import build
from src.CalendarAPI.genCredentials import authorization

tokenFile = str(THIS_FOLDER) + "/src/Authentication/Calendar/token.pickle"
class google_calendar:
    def __init__(self, calendarId):
        self.calendarId = calendarId
        self.shiften = shiften
        self.color = color
    #prepEventData needs shiftenExcel and datesExcel to create a list, that can be used to create events.
    #Only used for createEvents method.
    def __prepEventData__(self):
        data = []
        for i in range(len(self.datesExcel)):
            #Shiften (from configuration file)
            shiftEvent = self.shiften[self.shiftenExcel[i]]
            dateEvent = self.datesExcel[i]
            nameEvent = self.shiftenExcel[i]
        
            startHour = shiftEvent['startUur']
            endHour = shiftEvent['eindUur']
            colorID = self.color[shiftEvent['kleur']]  
            #convertTo_isoformat (from functions lib)
            dateTimes = convertTo_isoformat(dateEvent[2], months[dateEvent[1]], dateEvent[0], startHour, endHour)               
            data.append([self.shiftenExcel[i], locatie, colorID, dateTimes[0], dateTimes[1]])        
        return data
    def createEvents(self, datesExcel, shiftenExcel, beschrijving):
        self.datesExcel = datesExcel
        self.shiftenExcel = shiftenExcel
        self.data = self.__prepEventData__()
        events = []
        for i in range(len(self.data)):
            events.append(createEvent(self.data[i][0], self.data[i][1], self.data[i][2], self.data[i][3], self.data[i][4], beschrijving))
        self.events = events
        return self.events
    def writeEventsToCalendar(self, datesExcel, shiftenExcel):
        self.datesExcel = datesExcel
        self.shiftenExcel = shiftenExcel
        self.createEvents(self.datesExcel, self.shiftenExcel)
        self.writeCycleToCalendar()
    def getEvents(self):
        credentials = pickle.load(open(tokenFile, "rb"))
        service = build("calendar", "v3", credentials = credentials)
        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        print(now)
        print('Getting the upcoming 10 events')
        events_result = service.events().list(calendarId= self.calendarId, timeMin=now,
                                            maxResults=10, singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
        return events
    def createCycle(self, name, startDate, endDate, startCycle, period):
        date_startDate = date(startCycle[0], startCycle[1], startCycle[2])
        date_endDate = date(endDate[0], endDate[1], endDate[2])
        startDate = date(startDate[0], startDate[1], startDate[2])
        dates = []
        shiftenExcel = []
        date_1 = date_startDate
        d = math.ceil((abs(date_endDate-date_startDate).days+1)/period)
        for i in range(d):
            if date_1 >= startDate:
                dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
            date_1 += timedelta(days=period)
            shiftenExcel.append(name)
        return dates, shiftenExcel
    def create_monthCycle(self, name, startDate, endDate, startCycle, period, extend_period):
        date_startDate = date(startCycle[0], startCycle[1], startCycle[2])
        date_endDate = date(endDate[0], endDate[1], endDate[2])
        startDate = date(startDate[0], startDate[1], startDate[2])
        dates = []
        shiftenExcel = []
        date_1 = date_startDate
        d = math.ceil((abs(date_endDate-date_startDate).days+1)/period)
        prev_month = ""
        date_1 = startDate
        for i in range(d):
                if prev_month != date_1.strftime("%m"):
                    dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
                    prev_month = date_1.strftime("%m")
                else:
                    date_1 += timedelta(days=extend_period)
                    dates.append([int(date_1.strftime("%d")), maanden[date_1.strftime("%m")],date_1.strftime("%Y")])
                date_1 += timedelta(days=period)
                shiftenExcel.append(name)        
        return dates, shiftenExcel
    def writeCycleToCalendar(self):
        #If there's no token, it will be created bij the authorization function
        print("Token bein checked...\n")
        token = os.path.exists(tokenFile)
        print("Token present = ", token)
        if token == False:
            credentials = authorization()
            pickle.dump(credentials, open(tokenFile, "wb"))  

        credentials = pickle.load(open(tokenFile, "rb"))
        service = build("calendar", "v3", credentials = credentials)
        
        for i in range(len(self.events)):
            event = service.events().insert(calendarId = self.calendarId, body=self.events[i]).execute()
        return    
        
