from __future__ import print_function
# -*- coding:utf-8 -*-
__author__ = "ganbin"
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import re
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os
import datetime
import sheets_sample




try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = [
    'https://mail.google.com/',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.settings.basic',
    ]
CLIENT_SECRET_FILE = 'client_gmail_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'weekly_auto.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
service = discovery.build('gmail', 'v1', http=http)
def main():
    # threads = list_threads_with_label_query(service, label_ids=label_id, query='before:2017/09/30 after:2017/08/01')
    # subjects = list_subject(service, threads)
    # print(subjects)
    message = create_message_with_attachment('xiaoxi@nonda.us', 'ganbinwen@nonda.me')
    print(message)
    send_message(service, message)
    return
    values = sheets_sample.list_value()
    print(len(values))
    for v in values:
        email = v[1]
        if email:
            print(email)
            _count_daily_report(service, email)


def get_label_id(service, labelname='Daily Report'):
    labels = service.users().labels().list(userId='me').execute()
    label_id = ''  # if haven't got label_id , it's mean all labels.
    if labels.get('labels'):
        for label in labels['labels']:
            if labelname in label['name']:
                label_id = label['id']
                break
    return label_id


def list_threads_with_label_query(service, label_ids, query='', user_id='me'):
    response = service.users().threads().list(userId=user_id,
                                              labelIds=label_ids,
                                              q=query).execute()
    threads = []
    if 'threads' in response:
        threads.extend(response['threads'])

    while 'nextPageToken' in response:
        page_token = response['nextPageToken']
        response = service.users().threads().list(userId=user_id,
                                                  labelIds=label_ids,
                                                  pageToken=page_token,
                                                  q=query).execute()
        threads.extend(response['threads'])
    return threads


def list_subject(service, threads, _format='metadata'):
    subjects = []
    for thread in threads:
        thread_id = thread['id']
        metadata = service.users().threads().get(userId='me',
                                                 id=thread_id,
                                                 format=_format).execute()
        payloads = metadata['messages'][0]['payload']
        for payload in payloads['headers']:
            if payload['name'] == 'Subject':
                subjects.append(payload['value'])
                break
    return subjects




def create_message_with_attachment(sender, to):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = '【提醒】新人周得分填写'
    message['htmlBody'] = '请在16点前更新完毕项目进度情况，未及时更新要请大家下午茶哟<div class="gmail_chip gmail_drive_chip" style="width:396px;height:18px;max-height:18px;background-color:rgb(245,245,245);padding:5px;color:rgb(34,34,34);font-family:arial;font-style:normal;font-weight:bold;font-size:13px;border:1px solid rgb(221,221,221);line-height:1">' \
                          '附件为所有项目的具体信息。<br>' \
                          '本周CN team各个项目进度已按新版本更新。<br><br>'
    message_text = '请在16点前更新完毕项目进度情况，未及时更新要请大家下午茶哟<div class="gmail_chip gmail_drive_chip" style="width:396px;height:18px;max-height:18px;background-color:rgb(245,245,245);padding:5px;color:rgb(34,34,34);font-family:arial;font-style:normal;font-weight:bold;font-size:13px;border:1px solid rgb(221,221,221);line-height:1">' \
                          '附件为所有项目的具体信息。<br>' \
                          '本周CN team各个项目进度已按新版本更新。<br><br>'
    # message_text = 'Hi Managers，\n请尽早和新人进行上周目标及结果的1-1沟通，并于今天5:00PM前更新新人上周的得分。'

    msg = MIMEText(message_text)
    msg.add_header('mixed/alternative', _value='html')
    message.attach(msg)

    # msg = MIMEText(
    #     'https://drive.google.com/drive/u/1/folders/0B2TOYKCbyuaRLWFnZUw2RVV3OXc', _subtype='text')
    # msg.add_header('Content-Disposition', 'attachment', filename='附件')
    # message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}

# def create_message(sender, to):
#     message_text = 'Hi Managers，\n请尽早和新人进行上周目标及结果的1-1沟通，并于今天5:00PM前更新新人上周的得分。\n' \
#                    'https://drive.google.com/drive/u/1/folders/0B2TOYKCbyuaRLWFnZUw2RVV3OXc'
#     message = MIMEText(message_text)
#     message['to'] = to
#     message['from'] = sender
#     message['subject'] = '【提醒】新人周得分填写'
#     return {
#         'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()
#     }

def send_message(service, message, user_id='me'):
    message = (service.users().messages().send(userId=user_id,body=message)
           .execute())
    print('Message Id: %s' % message['id'])
    return message


def get_search_time():
    now = datetime.datetime.now()
    interval_days = 9
    last_monday = now - datetime.timedelta(days=interval_days)
    tomorrow = now - datetime.timedelta(days=1)
    if last_monday.weekday() == 0:  # 0 is Monday
        return last_monday, tomorrow
    print("Today is not Monday, May be have some errors.")


def _count_daily_report(service, email, label_ids='Label_21'):
    strf_time = '%Y/%m/%d'
    # 'before:2017/09/30 after:2017/08/01'
    last_monday, tomorrow = get_search_time()
    after_date = last_monday.strftime(strf_time)
    before_date = tomorrow.strftime(strf_time)
    filter_after = 'after:%s' % after_date
    filter_before = 'before:%s' % before_date
    filter_from = 'from:%s' % email
    match_days = []
    interval_day = datetime.timedelta(days=1)
    begin_day = last_monday
    over_day = last_monday + datetime.timedelta(days=5)
    while begin_day != over_day:
        match_days.append(begin_day)
        begin_day += interval_day
    lambda_date = lambda x: '%s/%s' %(x.month, x.day)
    match_date = [lambda_date(x) for x in match_days]

    print(filter_after, filter_before, match_date)
    query = '%s %s %s' %(filter_from, filter_after, filter_before)

    threads = list_threads_with_label_query(service, label_ids, query)
    subjects = list_subject(service, threads)
    print(subjects)
    all_subject = ''.join(subjects)
    all_subject = re.sub('[\D]','/',all_subject)  # only reserve digit
    re_date = re.compile(r'(\d{1,2}/{0,1}\d{1,2})')
    result = re_date.findall(all_subject)
    print(all_subject, result)
    weekly_report = []
    for d in match_date:
        if d in result:
            weekly_report.append((d, True))
        else:
            weekly_report.append((d, False))
    print(weekly_report)
    return weekly_report


def count_daily_report(email):
    return _count_daily_report(service, email)



if __name__ == '__main__':
    # count_daily_report(email='')
    main()
    # _count_daily_report(1,1)