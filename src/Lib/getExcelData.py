import xlrd, xlwt

def getDates(path_file):
    wb = xlrd.open_workbook(path_file) 
    sheet = wb.sheet_by_index(0) 
    row = []
    i = 0
    try:
        while sheet.cell_value(2,i) != "":   
            try: 
                int(sheet.cell_value(2,i))
                string = sheet.cell_value(0,i).lower()
                string = string.split()
                string.insert(0, int(sheet.cell_value(2,i)))
                row.append(string)
                i+=1
            except:
                i+=1
    except:
        pass
    return row

def getShiften(path_file):
    wb = xlrd.open_workbook(path_file) 
    sheet = wb.sheet_by_index(0) 
    i = 0
    try:
        while True:
            if sheet.cell_value(i,0).lower() == ("nuyts christine" or "christine nuyts"):
                x = i
                break
            i+=1
    except:
        pass
    try:
        i = 0
        row = []
        while True:
            row.append(sheet.cell_value(x,i).lower())
            i+=1
    except:
        pass
    for i in range(len(row)):
        if ":" in row[i]:
            x = i+1
            break
    del row[0:x]
    return row


    