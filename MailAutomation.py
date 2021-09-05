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
FOLDER_ROOT = 1

# Reference: https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat
# 1: plain, 2: HTML, 3: richtext
BODY_FORMAT = 3

SUBJECT_SCHEDULE_TAG = '【在宅勤務予定】'
SUBJECT_WORKSTART_TAG = '【在宅勤務開始】'
SUBJECT_WORKEND_TAG = '【在宅勤務終了】'

BODY_TITLE_OF_HONOR = 'さん\r\n'
BODY_DESU = 'です。'
BODY_SCHEDULE = 'に下記予定で在宅勤務いたします。\r\n'
BODY_WORKSTARTS = 'です。\r\n\r\n本日在宅勤務開始します。\r\n'
BODY_WORKENDS = 'です。\r\n\r\n本日在宅勤務終了します。\r\n'
BODY_BORDER = '------------------------------------------------------------------'
BODY_SIGNOFF = '以上、よろしくお願いいたします。\r\n'

# String used for locating reply mail body.
# Considering user signature may also include undersocres so use two lines to locate.
# 45 underscores
START_OF_REPLY_MAIL_BODY = '_____________________________________________\r\nFrom:'

class Configuration:
    config_file_name = 'config.json'

    # Customizable variable
    to_address = ''
    cc_address = ''
    my_name = ''
    supervisor_name = ''
    target_folder_name = ''
    time_delta = ''

class Outlook:
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    mapi_namespace = outlook_app.GetNamespace("MAPI")

    # Get all calendar items
    calender_items = mapi_namespace.GetDefaultFolder(FOLDER_CALENDAR).Items

    # Set sentmail(folder)
    sentmail = mapi_namespace.GetDefaultFolder(FOLDER_SENTMAIL)

    # Set Inbox
    inbox = mapi_namespace.GetDefaultFolder(FOLDER_INBOX)

    # Set root folder
    root_folder = mapi_namespace.Folders.Item(FOLDER_ROOT)

    # Target folder is not available by default
    target_folder = None

def get_configurations():
    # application is a frozen exe
    if getattr(sys, 'frozen', False):
        app_path = os.path.dirname(sys.executable)
    # application is a script file
    else:
        app_path = os.path.dirname(__file__)

    config_file_path = os.path.join(app_path, Configuration.config_file_name)

    # Load customizable variable from configuration file
    with open(config_file_path, encoding='utf-8') as config_file:
        config_dict = json.load(config_file)
        Configuration.to_address = config_dict['To']
        Configuration.cc_address = config_dict['Cc']
        Configuration.my_name = config_dict['MyName']
        Configuration.supervisor_name = config_dict['SupervisorName']
        Configuration.target_folder_name = config_dict['FolderName']
        Configuration.time_delta = int(config_dict['TimeDelta'])

def traverse_folder(par_parent_folder):
    try:
        Outlook.target_folder = par_parent_folder.Folders[Configuration.target_folder_name]
    except:
        for subfolder in par_parent_folder.Folders:
            traverse_folder(subfolder)

def send_schedule():
    # 勤務予定日は翌日なので、翌日の日付を取得
    local_work_date = datetime.today().date() + timedelta(Configuration.time_delta)      # Format: yyyy-mm-dd
    local_start_time = time.fromisoformat(START_TIME_STR)
    local_start_datetime = datetime.combine(local_work_date, local_start_time)
    local_end_time = time.fromisoformat(END_TIME_STR)
    local_end_datetime = datetime.combine(local_work_date, local_end_time)
    local_work_date_mm_dd = local_work_date.strftime("%#m/%#d")                          # Format: mm-dd without leading zero. Add a # between the % and the letter to remove leading zero.

    local_cal_items = Outlook.calender_items
    local_cal_items.IncludeRecurrences = True
    local_cal_items.Sort("[Start]")

    local_restriction = "[Start] >= '" + local_start_datetime.strftime("%Y-%m-%d %H:%M") + "' And [End] <= '" + local_end_datetime.strftime("%Y-%m-%d %H:%M") + "'"
    local_cal_items = local_cal_items.Restrict(local_restriction)

    # Make mail body
    local_body_list = []
    local_body_list.append(Configuration.supervisor_name + BODY_TITLE_OF_HONOR)
    local_body_list.append(Configuration.my_name + BODY_DESU)
    local_body_list.append('\r\n' + local_work_date_mm_dd + BODY_SCHEDULE)
    local_body_list.append(BODY_BORDER)
    for tmp_item in local_cal_items:
        tmp_subject = tmp_item.Subject
        tmp_time_str = "{0}～{1}".format(tmp_item.start.strftime("%H:%M"), tmp_item.end.strftime("%H:%M"))
        local_body_list.append(tmp_time_str + ' ' + tmp_subject)

    local_body_list.append(BODY_BORDER + '\r\n')
    local_body_list.append(BODY_SIGNOFF)

    local_mailbody = '\r\n'.join(local_body_list)

    local_new_mail = Outlook.outlook_app.CreateItem(0)
    local_new_mail.BodyFormat = BODY_FORMAT
    local_new_mail.To = Configuration.to_address
    local_new_mail.CC = Configuration.cc_address
    local_new_mail.Subject = SUBJECT_SCHEDULE_TAG + Configuration.my_name + ' ' + local_work_date_mm_dd
    local_new_mail.Body = local_mailbody
    local_new_mail.Display()
    print('メールを作成しました。')

def reply_mail(par_tag_for_search, par_tag_for_title, par_text_for_body):
    print('メール検索中。。。')
    local_is_found = False
    local_reply_mail = None
    local_body_list = []
    local_body_string = ''
    local_body_without_signature = ''

    # 当日の連絡なので、当日の日付を取得
    local_work_date = datetime.today().date()
    local_work_date_mm_dd = local_work_date.strftime("%#m/%#d")
    local_subject_to_find = par_tag_for_search + Configuration.my_name + ' ' + local_work_date_mm_dd

    local_sent_items = Outlook.sentmail.Items
    # 最新の送信メールから探す
    local_sent_items.Sort('[SentOn]', True)

    local_received_items = Outlook.target_folder.Items
    # 最新の受信メールから探す
    local_received_items.Sort('[ReceivedTime]', True)

    # 返信用のメールを探す
    for tmp_sent_item in local_sent_items:
        if local_subject_to_find in tmp_sent_item.Subject:
            local_is_found = True
            local_reply_mail = tmp_sent_item.Reply()

            local_body_string = local_reply_mail.Body
            # Get rid of user signature
            # Locate the start of reply mail and get all the string from start to the end.
            local_body_without_signature = local_body_string[local_body_string.index(START_OF_REPLY_MAIL_BODY):]
            # Replace original mail body with a non-signature version
            local_reply_mail.Body = local_body_without_signature

            for tmp_received_item in local_received_items:
                if local_subject_to_find in tmp_received_item.Subject:
                    if tmp_received_item.ReceivedTime > tmp_sent_item.SentOn:
                        local_reply_mail = tmp_received_item.Reply()

                    break

            break

    if local_is_found == True:
        local_reply_mail.BodyFormat = BODY_FORMAT
        local_reply_mail.Subject = par_tag_for_title + Configuration.my_name + ' ' + local_work_date_mm_dd
        local_body_list.append(Configuration.supervisor_name + BODY_TITLE_OF_HONOR)
        local_body_list.append(Configuration.my_name + par_text_for_body)
        local_body_list.append(BODY_SIGNOFF)
        local_reply_mail.Body = '\r\n'.join(local_body_list) + local_reply_mail.Body
        local_reply_mail.To = Configuration.to_address
        local_reply_mail.CC = Configuration.cc_address
        local_reply_mail.Display()
        print('メールを作成しました。')
    else:
        print(local_subject_to_find + ' のメールが見つかりません。')

if __name__ == "__main__":
    try:
        function_selection = int(input('機能をご選択ください(1:予定連絡、2:開始連絡、3:終了連絡)：'))
    except:
        print('全角/半角数字1、2または3をご入力ください。')

    try:
        get_configurations()
        traverse_folder(Outlook.root_folder)

        # 予定連絡：翌日の予定を上司に送付する
        if function_selection == 1:
            send_schedule()

        # 開始連絡：本日の勤務開始連絡を上司に送付する
        elif function_selection == 2:
            reply_mail(SUBJECT_SCHEDULE_TAG, SUBJECT_WORKSTART_TAG, BODY_WORKSTARTS)

        # 終了連絡：本日の勤務終了連絡を上司に送付する
        elif function_selection == 3:
            reply_mail(SUBJECT_WORKSTART_TAG, SUBJECT_WORKEND_TAG, BODY_WORKENDS)

        # Unexpected input
        else:
            print('全角/半角数字1、2または3をご入力ください。')

    except Exception as e:
        print(e)

os.system('pause')