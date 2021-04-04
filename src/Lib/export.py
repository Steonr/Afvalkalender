import xlwt 
from xlwt import Workbook 
import os, inspect, sys
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
oneUp = os.path.dirname(currentdir)
#Append the main directory to system path --> otherwise importing the modules is not possible
sys.path.append(oneUp)

def writeToExcel(data, filename, oneCell):
    loc = "./Export/{}.xls".format(filename)
    # Workbook is created 
    wb = Workbook()     
    # add_sheet is used to create sheet. 
    sheet1 = wb.add_sheet('Sheet 1')     
    if len(data) == 0 or oneCell == True: #Onecell: Boolean that determines if data has to been written in one cell
        sheet1.write(0, 0, str(data))
    else:
        for i in range(len(data)):
            sheet1.write(i, 0, str(data[i]))   
    wb.save(loc) 
    return 0

import json
def writeToJson(data, filename):
    loc = "./Export/{}.json".format(filename)
    with open(loc, 'w') as outfile:
        json.dump(data, outfile)
    