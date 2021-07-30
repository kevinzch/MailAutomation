#!/usr/bin/env python
# -*- coding: utf-8 -*-

import win32com.client
from datetime import datetime, timedelta, time
import os
import sys
import json

START_TIME_STR = '06:00:00'
END_TIME_STR   = '22:00:00'

# Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
FOLDER_CALENDAR = 9
FOLDER_SENTMAIL = 5

# Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat
# 1: plain, 2: HTML, 3: richtext
BODY_FORMAT = 3      

MAIL_SUBJECT_TAG = '【在宅勤務予定】'
WORKSTART_TAG = '【在宅勤務開始連絡】'

GREETING_STR = '〇〇〇〇〇\n'
MORNING_GREETING_STR = 'おはようございます。'
WORKSTART_TEXT = '本日在宅勤務開始します。'
MAIL_BODY_BORDER = '------------------------------------------------------------------'
MAIL_BODY_SIGNOFF = '\n以上、よろしくお願いいたします。'

class Configration:
    configFileName = 'config.json'
    toAddr = ''
    ccAddr = ''
    selfName = ''
    supervisorName = ''

    # application is a frozen exe
    if getattr(sys, 'frozen', False):
        appPath = os.path.dirname(sys.executable)
    # application is a script file
    else:
        appPath = os.path.dirname(__file__)

    configFilePath = os.path.join(appPath, configFileName)    

class Outlook:
    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calender = namespace.GetDefaultFolder(FOLDER_CALENDAR).Items
    sentMail = namespace.GetDefaultFolder(FOLDER_SENTMAIL).Items

def getCalendarItems(start, end):
    tmpCalItems = Outlook.calender
    tmpCalItems.IncludeRecurrences = True
    tmpCalItems.Sort("[Start]")

    tmpRestriction = "[Start] >= '" + start.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + end.strftime("%Y-%m-%d %H:%M") + "'"
    tmpCalItems = tmpCalItems.Restrict(tmpRestriction)

    return tmpCalItems

def sendSchedule(mBody, date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    tmpMail = outlook.CreateItem(0)
    tmpMail.BodyFormat = BODY_FORMAT
    tmpMail.To = Configration.toAddr
    tmpMail.CC = Configration.ccAddr
    tmpMail.Subject = MAIL_SUBJECT_TAG + Configration.selfName + " " + date.strftime("%m/%d")
    tmpMail.Body = mBody
    tmpMail.Display()

def makeBodyForNewMail(calItems):
    tmpBodyList = []
    tmpBodyList.append(GREETING_STR)
    tmpBodyList.append(MAIL_BODY_BORDER)
    for item in calItems:
        subjectStr = item.subject
        timeStr = "{0}～{1}".format(item.start.strftime("%H:%M"), item.end.strftime("%H:%M"))
        tmpBodyList.append(subjectStr + ' ' + timeStr)

    tmpBodyList.append(MAIL_BODY_BORDER)
    tmpBodyList.append(MAIL_BODY_SIGNOFF)

    mBody = '\n'.join(tmpBodyList)
    return mBody

def getConfigration(filePath):
    with open(filePath, encoding='utf-8') as configFile:
        configDict = json.load(configFile)

        Configration.toAddr = configDict['To']
        Configration.ccAddr = configDict['Cc']
        Configration.selfName = configDict['SelfName']
        Configration.supervisorName = configDict['SupervisorName']

def sendWorkStartMail():
    # 本日の開始連絡なので、当日の日付を取得
    workDate = datetime.today().date()
    subjectToFind = MAIL_SUBJECT_TAG + Configration.selfName + " " + workDate.strftime("%m/%d")
    sentMailItems = Outlook.sentMail

    isMailFound = False

    for item in sentMailItems:
        if item.Subject == subjectToFind:
            replyMail = item.Reply()
            replyMail.Subject = WORKSTART_TAG + Configration.selfName + " " + workDate.strftime("%m/%d") + ' 8:15～'
            tmpBodyList = []
            tmpBodyList.append(Configration.supervisorName + 'さん\n')
            tmpBodyList.append(MORNING_GREETING_STR)
            tmpBodyList.append(Configration.selfName + 'です。\n')
            tmpBodyList.append(WORKSTART_TEXT)
            tmpBodyList.append(MAIL_BODY_SIGNOFF)
            replyMail.Body = '\n'.join(tmpBodyList) + replyMail.Body
            replyMail.To = Configration.toAddr
            replyMail.CC = Configration.ccAddr
            replyMail.Display()
            isMailFound = True

    if isMailFound == False:
        print('在宅勤務予定の送信済みメールが見つかりません。')

if __name__ == "__main__":
    # try:
        functionSel = int(input('機能をご選択ください(1:予定連絡、2:開始連絡、3:終了連絡)：'))
        getConfigration(Configration.configFilePath)

        # 予定連絡：翌日の予定を上司に送付する
        if functionSel == 1:
            # 勤務予定日は翌日なので、翌日の日付を取得
            workDate = datetime.today().date() + timedelta(1)
            startTime = time.fromisoformat(START_TIME_STR)
            startDateTime = datetime.combine(workDate, startTime)
            endTime = time.fromisoformat(END_TIME_STR)
            endDateTime = datetime.combine(workDate, endTime)
            
            calenderItems = getCalendarItems(startDateTime, endDateTime)
            mailBody = makeBodyForNewMail(calenderItems)
            sendSchedule(mailBody, workDate)

        # 開始連絡：本日の勤務開始の連絡を上司に送付する
        elif functionSel == 2:
            sendWorkStartMail()

        # 終了連絡
        elif functionSel == 3:
            pass

        # Unexpected input
        else:
            print('数字1、2または3をご入力ください。')
            
    # except:
    #     print('数字1、2または3をご入力ください。')

# os.system('pause')