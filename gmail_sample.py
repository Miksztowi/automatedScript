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


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    label_id = get_label_id(service)
    # threads = list_threads_with_label_query(service, label_ids=label_id, query='before:2017/09/30 after:2017/08/01')
    # subjects = list_subject(service, threads)
    # print(subjects)
    # message = create_message_with_attachment('xiaoxi@nonda.us', 'ganbinwen@nonda.me')
    # print(message)
    # send_message(service, message)




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


def _count_daily_report():
    pass


def create_message_with_attachment(sender, to):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = '【提醒】新人周得分填写'
    message_text = 'Hi Managers，\n请尽早和新人进行上周目标及结果的1-1沟通，并于今天5:00PM前更新新人上周的得分。'

    msg = MIMEText(message_text)
    message.attach(msg)

    msg = MIMEText(
        'https://drive.google.com/drive/u/1/folders/0B2TOYKCbyuaRLWFnZUw2RVV3OXc', _subtype='text')
    msg.add_header('Content-Disposition', 'attachment', filename='附件')
    message.attach(msg)

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
    message = (service.users().messages().send(userId=user_id, body=message)
           .execute())
    print('Message Id: %s' % message['id'])
    return message


def get_search_time():
    now = datetime.datetime.now()
    interval_days = 9
    last_monday = now - datetime.timedelta(days=interval_days)
    tomorrow = now - datetime.timedelta(days=1)
    if last_monday.weekday() == 0:  # 0 is Monday
        after_date = '%s/%s' % (last_monday.month, last_monday.day)
        before_date = '%s/%s' % (tomorrow.month, tomorrow.day)
        return after_date, before_date
    print("Today is not Monday, May be have some errors.")


if __name__ == '__main__':
    # main()
    print(get_time())
