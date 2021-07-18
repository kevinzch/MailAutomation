#!/usr/bin/env python
# -*- coding: utf-8 -*-

import win32com.client
from datetime import datetime, timedelta, time
import csv
import os

START_TIME_STR = '06:00:00'
END_TIME_STR   = '22:00:00'

FOLDER_CALENDAR = 9  # Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
BODY_FORMAT = 3      # 1: plain, 2: HTML, 3: richtext Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat

MAIL_SUBJECT_TAG = '【〇〇〇〇連絡】'

MAIL_BODY_GREETING = '〇〇〇〇〇\n'
MAIL_BODY_BORDER = '------------------------------------------------------------------'
MAIL_BODY_SIGNOFF = '\n〇〇〇〇〇'

class Settings:
    settingFilePath = os.path.join(os.path.dirname(__file__), 'settings.csv')

    toList = []
    ccList = []
    selfName = ''
    supervisorName = ''

def getCalendarItems(start, end):
    outlookNamespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    tmpCalItems = outlookNamespace.GetDefaultFolder(FOLDER_CALENDAR).Items
    tmpCalItems.IncludeRecurrences = True
    tmpCalItems.Sort("[Start]")

    tmpRestriction = "[Start] >= '" + start.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + end.strftime("%Y-%m-%d %H:%M") + "'"
    tmpCalItems = tmpCalItems.Restrict(tmpRestriction)

    return tmpCalItems

def sendSchedule(mBody, date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    tmpMail = outlook.CreateItem(0)
    tmpMail.BodyFormat = BODY_FORMAT
    tmpMail.To = ';'.join(Settings.toList)
    tmpMail.CC = ';'.join(Settings.ccList)
    tmpMail.Subject = MAIL_SUBJECT_TAG + str(Settings.selfName) + " " + date.strftime("%m/%d")
    tmpMail.Body = mBody
    tmpMail.Display()

def makeMailBody(calItems):
    tmpBodyList = []
    tmpBodyList.append(MAIL_BODY_GREETING)
    tmpBodyList.append(MAIL_BODY_BORDER)
    for item in calItems:
        subjectStr = item.subject
        timeStr = "{0}～{1}".format(item.start.strftime("%H:%M"), item.end.strftime("%H:%M"))
        tmpBodyList.append(subjectStr + ' ' + timeStr)

    tmpBodyList.append(MAIL_BODY_BORDER)
    tmpBodyList.append(MAIL_BODY_SIGNOFF)

    mBody = '\n'.join(tmpBodyList)
    return mBody

def getSettings(filePath):
    with open(filePath, encoding='utf-8') as tmpFile:
        tmpData = csv.DictReader(tmpFile, delimiter=',')
        tmpDict = {}
        for row in tmpData:
            for key, value in row.items():
                if value is not None and value != '':
                    try:
                        tmpDict[key].append(value)
                    except KeyError:
                        tmpDict[key] = [value]
        
        Settings.toList = list(tmpDict['To'])
        Settings.ccList = list(tmpDict['CC'])
        Settings.selfName = str(tmpDict['SelfName'])
        Settings.supervisorName = str(tmpDict['SupervisorName'])

workDate = datetime.today().date() + timedelta(1)
startTime = time.fromisoformat(START_TIME_STR)
startDateTime = datetime.combine(workDate, startTime)
endTime = time.fromisoformat(END_TIME_STR)
endDateTime = datetime.combine(workDate, endTime)

getSettings(Settings.settingFilePath)
print(str(Settings.selfName))
calenderItems = getCalendarItems(startDateTime, endDateTime)
mailBody = makeMailBody(calenderItems)
sendSchedule(mailBody, workDate)