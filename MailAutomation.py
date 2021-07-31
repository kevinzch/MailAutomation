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

SUBJECT_SCHEDULE_TAG = '【在宅勤務予定】'
SUBJECT_WORKSTART_TAG = '【在宅勤務開始】'
SUBJECT_WORKEND_TAG = '【在宅勤務終了】'

BODY_PERSONAL_TITLE = 'さん\n'
BODY_SCHEDULE = 'です。\n明日下記予定で在宅勤務いたします。\n'
BODY_WORKSTART = 'です。\n本日在宅勤務開始します。\n'
BODY_BORDER = '------------------------------------------------------------------'
BODY_SIGNOFF = '以上、よろしくお願いいたします。'

class Configration:
    configFileName = 'config.json'

    # application is a frozen exe
    if getattr(sys, 'frozen', False):
        appPath = os.path.dirname(sys.executable)
    # application is a script file
    else:
        appPath = os.path.dirname(__file__)

    configFilePath = os.path.join(appPath, configFileName)

    with open(configFilePath, encoding='utf-8') as configFile:
        configDict = json.load(configFile)
        toAddr = configDict['To']
        ccAddr = configDict['Cc']
        selfName = configDict['SelfName']
        supervisorName = configDict['SupervisorName']

class Outlook:
    outlookApp = win32com.client.Dispatch("Outlook.Application")
    namespace = outlookApp.GetNamespace("MAPI")
    calender = namespace.GetDefaultFolder(FOLDER_CALENDAR).Items
    sentMail = namespace.GetDefaultFolder(FOLDER_SENTMAIL).Items

def sendSchedule():
    # 勤務予定日は翌日なので、翌日の日付を取得
    tmpWorkDate = datetime.today().date() + timedelta(1)
    tmpStartTime = time.fromisoformat(START_TIME_STR)
    tmpStartDateTime = datetime.combine(tmpWorkDate, tmpStartTime)
    tmpEndTime = time.fromisoformat(END_TIME_STR)
    tmpEndDateTime = datetime.combine(tmpWorkDate, tmpEndTime)

    tmpCalItems = Outlook.calender
    tmpCalItems.IncludeRecurrences = True
    tmpCalItems.Sort("[Start]")

    tmpRestriction = "[Start] >= '" + tmpStartDateTime.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + tmpEndDateTime.strftime("%Y-%m-%d %H:%M") + "'"
    tmpCalItems = tmpCalItems.Restrict(tmpRestriction)

    tmpBodyList = []
    tmpBodyList.append(Configration.supervisorName + BODY_PERSONAL_TITLE)
    tmpBodyList.append(Configration.selfName + BODY_SCHEDULE)
    tmpBodyList.append(BODY_BORDER)
    for tmpItem in tmpCalItems:
        tmpSubjectStr = tmpItem.subject
        tmpTimeStr = "{0}～{1}".format(tmpItem.start.strftime("%H:%M"), tmpItem.end.strftime("%H:%M"))
        tmpBodyList.append(tmpSubjectStr + ' ' + tmpTimeStr)

    tmpBodyList.append(BODY_BORDER + '\n')
    tmpBodyList.append(BODY_SIGNOFF)

    tmpMailBody = '\n'.join(tmpBodyList)

    tmpNewMail = Outlook.outlookApp.CreateItem(0)
    tmpNewMail.BodyFormat = BODY_FORMAT
    tmpNewMail.To = Configration.toAddr
    tmpNewMail.CC = Configration.ccAddr
    tmpNewMail.Subject = SUBJECT_SCHEDULE_TAG + Configration.selfName + " " + tmpWorkDate.strftime("%m/%d")
    tmpNewMail.Body = tmpMailBody
    tmpNewMail.Display()

def sendWorkStartEndMail(parTagToSearch, parTagForTitle):
    # 当日の連絡なので、当日の日付を取得
    tmpWorkDate = datetime.today().date()
    tmpSubjectToFind = parTagToSearch + Configration.selfName + " " + tmpWorkDate.strftime("%m/%d")
    tmpSentMailItems = Outlook.sentMail

    tmpIsMailFound = False

    print(tmpSubjectToFind)

    for tmpItem in tmpSentMailItems:
        if tmpSubjectToFind in tmpItem.Subject:
            tmpReplyMail = tmpItem.Reply()
            tmpReplyMail.Subject = parTagForTitle + Configration.selfName + " " + tmpWorkDate.strftime("%m/%d")
            tmpBodyList = []
            tmpBodyList.append(Configration.supervisorName + BODY_PERSONAL_TITLE)
            tmpBodyList.append(Configration.selfName + BODY_WORKSTART)
            tmpBodyList.append(BODY_SIGNOFF)
            tmpReplyMail.Body = '\n'.join(tmpBodyList) + tmpReplyMail.Body
            tmpReplyMail.To = Configration.toAddr
            tmpReplyMail.CC = Configration.ccAddr
            tmpReplyMail.Display()
            tmpIsMailFound = True

    if tmpIsMailFound == False:
        print(parTagToSearch + 'のメールが見つかりません。')

if __name__ == "__main__":
    try:
        functionSel = int(input('機能をご選択ください(1:予定連絡、2:開始連絡、3:終了連絡)：'))

        # 予定連絡：翌日の予定を上司に送付する
        if functionSel == 1:            
            sendSchedule()

        # 開始連絡：本日の勤務開始連絡を上司に送付する
        elif functionSel == 2:
            sendWorkStartEndMail(SUBJECT_SCHEDULE_TAG, SUBJECT_WORKSTART_TAG)

        # 終了連絡：本日の勤務終了連絡を上司に送付する
        elif functionSel == 3:
            sendWorkStartEndMail(SUBJECT_WORKSTART_TAG, SUBJECT_WORKEND_TAG)

        # Unexpected input
        else:
            print('数字1、2または3をご入力ください。')

        print('メールを作成しました。')
    except:
        print('数字1、2または3をご入力ください。')

os.system('pause')