import sys, os, inspect
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(THIS_FOLDER)
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
    def createEvents(self, datesExcel, shiftenExcel):
        self.datesExcel = datesExcel
        self.shiftenExcel = shiftenExcel
        self.data = self.__prepEventData__()
        events = []
        for i in range(len(self.data)):
            events.append(createEvent(self.data[i][0], self.data[i][1], self.data[i][2], self.data[i][3], self.data[i][4], ""))
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
    #Generates cycle pattern. Can be used when there is a defined pattern to put in a calendar.
    def cycleGen(self):
        self.cycleList = []
        for i in range(7):
                shift = "Nacht"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for i in range(3):
                shift = "Vrij"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for i in range(7):
                shift = "Late"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for i in range(2):
                shift = "Vrij"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for i in range(7):
                shift = "Vroege"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for i in range(2):
                shift = "Vrij"
                self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        for j in range(4):
            for i in range(5):
                    shift = "Dag"
                    self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
            for i in range(2):
                    shift = "Vrij"
                    self.cycleList.append([shift, shiftenGerben[shift]["kleur"], shiftenGerben[shift]["hourStart"], shiftenGerben[shift]["hourEnd"]])
        return self.cycleList
    #cycle_createEvents doesn't use shiftenExcel and datesExcel
    def cycle_createEvents(self, startDate, endDate, beginCycle, tekst):
        self.cycleList = self.cycleGen()
        #Number of days
        d = abs(endDate-startDate).days+1
        #Number where cycle begins
        cycleNum = abs(startDate - beginCycle).days
        j = cycleNum
        if cycleNum > len(self.cycleList) - 1:
            #When cycle is used more than once --> check new cyclenumber
            j = cycleNum % len(self.cycleList)

        
        data = []
        events = []
        for i in range(d):
            data.append([])
            #Check if cycle has to run again:
            if j > len(self.cycleList)-1: 
                j = 0
            #Check the date of the shift:
            shiftDate = startDate + timedelta(days=i)
            
            #Shiftname:
            data[i].append(self.cycleList[j][0])
            #Shiftcolor:
            data[i].append(self.cycleList[j][1])
            #Start hour in RFC3339 format
            data[i].append(shiftDate.strftime("%Y-%m-%d")+'T'+self.cycleList[j][2]) 
            #End hour in RFC3339 format
            if self.cycleList[j][0] == "Nacht":
                    endshiftDate = startDate + timedelta(days=i+1)
                    data[i].append(endshiftDate.strftime("%Y-%m-%d")+'T'+self.cycleList[j][3])
            else:
                    data[i].append(shiftDate.strftime("%Y-%m-%d")+'T'+self.cycleList[j][3]) 
            #Make an event for the right shift --> google Calendar format
            events.append(createEvent(data[i][0], 'Exxonmobil, Meerhout', data[i][1], data[i][2], data[i][3], tekst))
            j +=1
        self.events = events
        return self.events

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
        
