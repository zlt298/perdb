import os
import xlrd
import madgetech.madgetech_2 as mt2
from madgetech.madgetech_2_tow import TOWprocessor
from madgetech.madgetech_2_cs import CorrosionSensorProcessor


def loadPERdict():
    try:
    #Open the PER Meta Data excel spreadsheet and create a dictionary with serial's as the key and
    #various meta data as the value.
        workbook = xlrd.open_workbook('..\PER Meta Data.xlsx')
        sheet = workbook.sheet_by_index(0)

        PERdict = {}

        for rx in [r for r in range(sheet.nrows) if r > 1]:
            r = sheet.row(rx)
            if r[0].value!='':
                serial = r[0].value
                site = r[1].value
                location = r[2].value
                project = r[3].value
                startDateTuple = xlrd.xldate.xldate_as_tuple(r[4].value,workbook.datemode)
                endDateTuple = xlrd.xldate.xldate_as_tuple(r[5].value,workbook.datemode)
                timeZone = r[6].value
                sensorType = r[7].value
                condition = r[8].value
                climate = r[9].value
                coord = (r[10].value,r[11].value)
                chl = r[12].value
                sul = r[13].value
                notes = r[14].value
                PERdict[serial] = [site,location,project,startDateTuple,endDateTuple,
                                   timeZone,sensorType,condition,climate,coord]
        return PERdict
    except Exception as e:
        print e
        raise Exception("Cannot find/open 'PER Meta Data.xlsx' in the parent directory, or the file format is faulty")

def inversePERdict():
    try:
    #Open the PER Meta Data Excel spreadsheet and create a dictionary with site + location as key and
    #data logger serial as the value. Used to determine which loggers are at each site
        
        workbook = xlrd.open_workbook('..\PER Meta Data.xlsx')
        sheet = workbook.sheet_by_index(0)

        invPER = {}

        for rx in [r for r in range(sheet.nrows) if r > 1]:
            r = sheet.row(rx)
            if r[0].value!='':
                serial = r[0].value
                site = r[1].value
                location = r[2].value
                if invPER.has_key(site + location):
                    invPER[site + location].append(serial)
                else:
                    invPER[site + location] = [serial]
        return invPER
    except Exception as e:
        print e
        raise Exception("Cannot find/open 'PER Meta Data.xlsx' in the parent directory, or the file format is faulty")

