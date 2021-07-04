#Python 3.8.3

import win32com.client
from datetime import datetime, timedelta, time

START_TIME_STR = "06:00:00"
END_TIME_STR   = "22:00:00"

FOLDER_CALENDAR = 9  # Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
FORMAT_HTML = 2      # Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat

MAIL_TO = "to@outlook.com"
MAIL_CC = "cc@outlook.com"
MAIL_SUBJECT_TAG = "【〇〇】"
MAIL_SUBJECT_NAME = "〇〇"

MAIL_BODY_GREETING = "〇〇〇〇〇\n"
MAIL_BODY_BORDER = "------------------------------------------------------------------"
MAIL_BODY_SIGNOFF = "\n〇〇〇〇〇"

def getCalendarItems(start, end):
    outlookNamespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calItems = outlookNamespace.GetDefaultFolder(FOLDER_CALENDAR).Items
    calItems.IncludeRecurrences = True
    calItems.Sort("[Start]")

    restriction = "[Start] >= '" + start.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + end.strftime("%Y-%m-%d %H:%M") + "'"
    calItems = calItems.Restrict(restriction)

    return calItems

def sendSchedule(mBody, date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    newMail = outlook.CreateItem(0)
    newMail.BodyFormat = FORMAT_HTML
    newMail.To = MAIL_TO
    newMail.CC = MAIL_CC
    newMail.Subject = MAIL_SUBJECT_TAG + MAIL_SUBJECT_NAME + " " + date.strftime("%m/%d")
    newMail.Body = mBody
    newMail.Display()

def makeMailBody(calItems):
    bodyList = []
    bodyList.append(MAIL_BODY_GREETING)
    bodyList.append(MAIL_BODY_BORDER)
    for item in calItems:
        subjectStr = item.subject
        timeStr = "{0}～{1}".format(item.start.strftime("%H:%M"), item.end.strftime("%H:%M"))
        bodyList.append(subjectStr + ' ' + timeStr)

    bodyList.append(MAIL_BODY_BORDER)
    bodyList.append(MAIL_BODY_SIGNOFF)

    mBody = '\n'.join(bodyList)
    return mBody

workDate = datetime.today().date() + timedelta(1)
startTime = time.fromisoformat(START_TIME_STR)
startDateTime = datetime.combine(workDate, startTime)
endTime = time.fromisoformat(END_TIME_STR)
endDateTime = datetime.combine(workDate, endTime)

calenderItems = getCalendarItems(startDateTime, endDateTime)
mailBody = makeMailBody(calenderItems)
sendSchedule(mailBody, workDate)