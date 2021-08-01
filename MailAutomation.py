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
FOLDER_INBOX = 6

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
    config_file_name = 'config.json'

    # application is a frozen exe
    if getattr(sys, 'frozen', False):
        app_path = os.path.dirname(sys.executable)
    # application is a script file
    else:
        app_path = os.path.dirname(__file__)

    config_file_path = os.path.join(app_path, config_file_name)

    with open(config_file_path, encoding='utf-8') as config_file:
        config_dict = json.load(config_file)
        to_address = config_dict['To']
        cc_address = config_dict['Cc']
        my_name = config_dict['MyName']
        supervisor_name = config_dict['SupervisorName']

class Outlook:
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    mapi_namespace = outlook_app.GetNamespace("MAPI")
    calender_items = mapi_namespace.GetDefaultFolder(FOLDER_CALENDAR).Items
    sent_items = mapi_namespace.GetDefaultFolder(FOLDER_SENTMAIL).Items
    received_items = mapi_namespace.GetDefaultFolder(FOLDER_INBOX).Items

def send_schedule():
    # 勤務予定日は翌日なので、翌日の日付を取得
    local_work_date = datetime.today().date() + timedelta(1)
    local_start_time = time.fromisoformat(START_TIME_STR)
    local_start_datetime = datetime.combine(local_work_date, local_start_time)
    local_end_time = time.fromisoformat(END_TIME_STR)
    local_end_datetime = datetime.combine(local_work_date, local_end_time)

    local_cal_items = Outlook.calender_items
    local_cal_items.IncludeRecurrences = True
    local_cal_items.Sort("[Start]")

    local_restriction = "[Start] >= '" + local_start_datetime.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + local_end_datetime.strftime("%Y-%m-%d %H:%M") + "'"
    local_cal_items = local_cal_items.Restrict(local_restriction)

    local_body_list = []
    local_body_list.append(Configration.supervisor_name + BODY_PERSONAL_TITLE)
    local_body_list.append(Configration.my_name + BODY_SCHEDULE)
    local_body_list.append(BODY_BORDER)
    for tmp_item in local_cal_items:
        tmp_subject = tmp_item.subject
        tmp_time_str = "{0}～{1}".format(tmp_item.start.strftime("%H:%M"), tmp_item.end.strftime("%H:%M"))
        local_body_list.append(tmp_subject + ' ' + tmp_time_str)

    local_body_list.append(BODY_BORDER + '\n')
    local_body_list.append(BODY_SIGNOFF)

    local_mailbody = '\n'.join(local_body_list)

    local_new_mail = Outlook.outlook_app.CreateItem(0)
    local_new_mail.BodyFormat = BODY_FORMAT
    local_new_mail.To = Configration.to_address
    local_new_mail.CC = Configration.cc_address
    local_new_mail.Subject = SUBJECT_SCHEDULE_TAG + Configration.my_name + ' ' + local_work_date.strftime("%m/%d")
    local_new_mail.Body = local_mailbody
    local_new_mail.Display()

def reply_mail(par_tag_for_search, par_tag_for_title):
    # 当日の連絡なので、当日の日付を取得
    local_work_date = datetime.today().date()
    local_subject_to_find = par_tag_for_search + Configration.my_name + ' ' + local_work_date.strftime("%m/%d")

    local_sent__items = Outlook.sent_items
    # 最新の送信メールから探す
    local_sent__items.Sort('[SentOn]', True)

    local_received_items = Outlook.received_items
    # 最新の受信メールから探す
    local_received_items.Sort('[ReceivedTime]', True)

    local_is_found = False

    print(local_subject_to_find)

    for tmp_item in local_sent__items:
        if local_subject_to_find in tmp_item.Subject:
            tmp_reply_mail = tmp_item.Reply()
            tmp_reply_mail.Subject = par_tag_for_title + Configration.my_name + ' ' + local_work_date.strftime("%m/%d")
            tmp_body_list = []
            tmp_body_list.append(Configration.supervisor_name + BODY_PERSONAL_TITLE)
            tmp_body_list.append(Configration.my_name + BODY_WORKSTART)
            tmp_body_list.append(BODY_SIGNOFF)
            tmp_reply_mail.Body = '\n'.join(tmp_body_list) + tmp_reply_mail.Body
            tmp_reply_mail.To = Configration.to_address
            tmp_reply_mail.CC = Configration.cc_address
            tmp_reply_mail.Display()
            local_is_found = True

            break

    if local_is_found == False:
        print(par_tag_for_search + 'のメールが見つかりません。')

if __name__ == "__main__":
    try:
        function_selection = int(input('機能をご選択ください(1:予定連絡、2:開始連絡、3:終了連絡)：'))

        # 予定連絡：翌日の予定を上司に送付する
        if function_selection == 1:
            send_schedule()

        # 開始連絡：本日の勤務開始連絡を上司に送付する
        elif function_selection == 2:
            reply_mail(SUBJECT_SCHEDULE_TAG, SUBJECT_WORKSTART_TAG)

        # 終了連絡：本日の勤務終了連絡を上司に送付する
        elif function_selection == 3:
            reply_mail(SUBJECT_WORKSTART_TAG, SUBJECT_WORKEND_TAG)

        # Unexpected input
        else:
            print('数字1、2または3をご入力ください。')

        print('メールを作成しました。')
    except:
        print('数字1、2または3をご入力ください。')

os.system('pause')