import os
import sys
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
sys.path.append(THIS_FOLDER)
def createEvent(shiftName, locatie, kleurCode, dateTimeBegin, dateTimeEnd, description):
    event = ( {
                'summary': shiftName,
                'location': locatie,
                'description': description,
                'colorId': kleurCode,
                'start': {
                'dateTime': dateTimeBegin,
                'timeZone': 'Europe/Brussels',
                        },
                'end': {
                'dateTime': dateTimeEnd,
                'timeZone': 'Europe/Brussels',
                        },
                })
    return event

#Add zero to number to get 01 format
def addZero(x):  
    if  int(x)<10:
        x = "0{}".format(x)
    return str(x)

def convertTo_isoformat(jaar, maand, dag, startUur, eindUur):
    maand = addZero(maand)
    dag = addZero(dag)
    strISOstart = "{}-{}-{}T{}:00.000".format(jaar, maand, dag, startUur)
    strISOend = "{}-{}-{}T{}:00.000".format(jaar, maand, dag, eindUur)
    return strISOstart, strISOend

from config import *
def prepEventData(datesExcel,shiftenExcel):
    data = []
    for i in range(len(datesExcel)):
        try:
            shiftEvent = shiften[shiftenExcel[i].upper()]
            dateEvent = datesExcel[i]
            nameEvent = shiftenExcel[i].upper()
        
            startHour = shiftEvent['startUur']
            endHour = shiftEvent['eindUur']
            colorID = color[shiftEvent['kleur']]  

            dateTimes = convertTo_isoformat(dateEvent[2], months[dateEvent[1]], dateEvent[0], startHour, endHour)               
            data.append([shiftenExcel[i].upper(), locatie, colorID, dateTimes[0], dateTimes[1]])        
        except:
            pass
    return data
from config import shiften
import datetime
import time
import numpy as np
class manData:
    def __init__(self, data):       
        self.data = data
        self.shiftenCount = []
        self.shiftenDuration = []
        
    def convDatetimeToHour(self, dateTime):
        conv = dateTime.total_seconds()
        conv = str(datetime.timedelta(seconds=conv))
        return conv[:-3]

    def calcTdelta_string(self, t1, t2, FMT):
        tdelta = datetime.datetime.strptime(t2, FMT) - datetime.datetime.strptime(t1, FMT) 
        tdelta = tdelta - datetime.timedelta(minutes = pauze)
        return tdelta

    def convHourToDatetime(self, hour, FMT):     
        return datetime.datetime.strptime(hour, FMT)

    def durationShift(self):        
        for key, values in shiften.items():            
            startHour = shiften[key]["startUur"]
            endHour = shiften[key]["eindUur"]
            tdelta = self.calcTdelta_string(startHour, endHour, "%H:%M")
            tdelta_string = self.convDatetimeToHour(tdelta)
            self.shiftenDuration.append([key, tdelta_string])
        return self.shiftenDuration

    def countShiften(self):
        for key, values in shiften.items():
            c = 0
            for i in range(len(self.data)):
                x = self.data[i].count(key)
                c +=x
            self.shiftenCount.append([key,c])       
        return self.shiftenCount

    def dataInfo(self):
        i = 0
        arr = []
        days = 0
        for key, values in shiften.items():
            if key != "V" and key !="VRIJ" and key !="ADV":
                x = datetime.datetime.strptime(self.durationShift()[i][1], "%H:%M").time()
                shiftcounter = self.countShiften()[i][1]
                hours = x.hour * shiftcounter
                minutes = x.minute * shiftcounter
                arr.append([key, days, hours, minutes])     
            i+=1    
        for i in range(len(arr)):
            arr[i][2] += int(arr[i][3]/60) 
            arr[i][3] = arr[i][3] % 60
            arr[i][1] = int(arr[i][2]/8)
            arr[i][2] = arr[i][2] % 8
        return arr # = Shift, Days, Hours, Minutes
    
    def countTotalHours(self):
        i = 0
        hourssum = 0
        minutessum = 0
        for key, values in shiften.items():
            if key != "V" and key !="VRIJ":# and key !="ADV":
                x = datetime.datetime.strptime(self.durationShift()[i][1], "%H:%M").time()
                shiftcounter = self.countShiften()[i][1]
                hours = x.hour * shiftcounter
                minutes = x.minute * shiftcounter
                hourssum += hours
                minutessum += minutes
            i+=1
        hourssum += int(minutessum/60)
        minutessum = minutessum % 60
        return hourssum, minutessum

    def kwaartaalInfo(self, dates):     
        #Count period of dates:
        beginDay =  int(dates[0][0])
        beginMonth = int(months[dates[0][1]])
        beginYear = int(dates[0][2])
        endDay =  int(dates[-1][0])
        endMonth = int(months[dates[-1][1]])
        endYear = int(dates[-1][2])
        tdelta = datetime.datetime(  year = endYear,    \
                                     month= endMonth,   \
                                     day= endDay)       \
                - datetime.datetime(  year = beginYear, \
                                     month= beginMonth, \
                                     day= beginDay)
        #Totaal aantal tijd voor één kwartaal:
        hours, minutes = self.countTotalHours()                     #Total working hours from data
        daysReal = tdelta.days                                      #Total days from dates  
        weeks = np.divide(daysReal, 7)                              #Total days from quarter of year
        weeks = round(weeks)
        totalDays = weeks*5
        totalHours = totalDays*werktijd[0] 
       
        totalMinutes = totalDays*werktijd[1] + totalHours*60        #Total minutes theoretical 
        minutes += hours*60                                         #Total minutes from data
        err = minutes-totalMinutes
    
        errHours = int(np.divide(abs(err), 60))
        errMinutes = abs(err) % 60
        if err < 0:
            x = -1
        else:
            x = 1

        workPer = np.divide(minutes, totalMinutes)
        workPer = round(workPer,4)*100
        return x, errHours, int(errMinutes), workPer                #Hours worked --> (+ = to much, - = to little)
import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
#Locate client_secretFile
client_secretFile = os.path.join(THIS_FOLDER, 'src/Authentication/Gmail/client_secret.json')
#Locate tokenFile
tokenFile = os.path.join(THIS_FOLDER, 'src/Authentication/Gmail/token.pickle')

def createCreds():
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(tokenFile):
        with open(tokenFile, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secretFile, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(tokenFile,'wb') as token:
            pickle.dump(creds, token)
    return creds