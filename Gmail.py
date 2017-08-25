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


import settings
import logging

# if retrive credentials, don't note these codes.
# try:
#     import argparse
#     flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
# except ImportError:
#     flags = None


class GamilAPI(object):
    def __init__(self):
        self._scopes = [
            'https://mail.google.com/',
            'https://www.googleapis.com/auth/gmail.modify',
            'https://www.googleapis.com/auth/gmail.readonly',
            'https://www.googleapis.com/auth/gmail.settings.basic',
            ]
        self._client_secret_file = 'client_gmail_secret.json'
        self._application_name = 'Gmail API Python Quickstart'
        self.service = self._get_service()
        self.label_id = self._get_label_id()

        logging.basicConfig(filename='AutoScripts.log', level='DEBUG')
        self.logger = logging.getLogger(__name__)

    def _get_credentials(self):
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'weekly_auto.json')
        store = Storage(credential_path)
        credentials = store.get()

        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self._client_secret_file, self._scopes)
            flow.user_agent = self._application_name
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def _get_service(self):
        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)
        return service

    def _get_label_id(self):
        labels = self.service.users().labels().list(userId='me').execute()
        label_id = ''  # if haven't got label_id , it's mean all labels.
        if labels.get('labels'):
            for label in labels['labels']:
                if settings.LABEL_NAME in label['name']:
                    label_id = label['id']
                    break
        return label_id

    def list_threads_with_label_query(self, label_ids, query='', user_id='me'):
        response = self.service.users().threads().list(userId=user_id,
                                                  labelIds=label_ids,
                                                  q=query).execute()
        threads = []
        if 'threads' in response:
            threads.extend(response['threads'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = self.service.users().threads().list(userId=user_id,
                                                      labelIds=label_ids,
                                                      pageToken=page_token,
                                                      q=query).execute()
            threads.extend(response['threads'])
        return threads

    def list_subject(self, threads, format='metadata'):
        subjects = []
        for thread in threads:
            thread_id = thread['id']
            metadata = self.service.users().threads().get(userId='me',
                                                     id=thread_id,
                                                     format=format).execute()
            payloads = metadata['messages'][0]['payload']
            for payload in payloads['headers']:
                if payload['name'] == 'Subject':
                    subjects.append(payload['value'])
                    break
        return subjects

    def create_message(self, sender, to, subject, text):
        message_text = text
        message = MIMEText(message_text, 'html')
        message['to'] = to
        message['from'] = sender
        message['subject'] = subject
        return {
            'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()
        }

    def send_message(self, message, user_id='me'):
        message = (self.service.users().messages().send(userId=user_id,body=message)
               .execute())
        self.logger.debug('Reminding email send successfully, Message Id: %s' % message['id'])
        return message

    def _get_search_date(self):
        now = datetime.datetime.now()
        interval_days = 4
        last_monday = now - datetime.timedelta(days=int(interval_days))
        tomorrow = now + datetime.timedelta(days=1)
        if last_monday.weekday() == 0:  # 0 is Monday
            return last_monday, tomorrow
        self.logger.error("Today is not Monday, if you persist in starting script, you must update settings.py")

    def get_match_date(self):
        last_monday, tomorrow = self._get_search_date()
        match_days = []
        interval_day = datetime.timedelta(days=1)
        begin_day = last_monday
        over_day = last_monday + datetime.timedelta(days=5)
        while begin_day != over_day:
            match_days.append(begin_day)
            begin_day += interval_day
        lambda_date = lambda x: '%s/%s' % (x.month, x.day)
        match_date = [lambda_date(x) for x in match_days]
        return match_date

    def _handle_query_with_date(self, email):
        last_monday, tomorrow = self._get_search_date()
        # 'before:2017/09/30 after:2017/08/01'
        strf_time = '%Y/%m/%d'
        after_date = last_monday.strftime(strf_time)
        before_date = tomorrow.strftime(strf_time)
        filter_after = 'after:%s' % after_date
        filter_before = 'before:%s' % before_date
        filter_from = 'from:%s' % email
        query = '%s %s %s' % (filter_from, filter_after, filter_before)
        return query

    def count_daily_with_email(self, email):
        query = self._handle_query_with_date(email)
        match_date = self.get_match_date()
        threads = self.list_threads_with_label_query(self.label_id, query)
        subjects = self.list_subject(threads)
        all_subject = ''.join(subjects)
        all_subject = re.sub('[\D]','/',all_subject)  # only reserve digit
        re_date = re.compile(settings.RE_DATE_RULE)
        result = re_date.findall(all_subject)
        weekly_report = []
        for d in match_date:
            find = False
            for r in result:
                if (r in d) or (d in r):
                    find = True
                    weekly_report.append((d, True))
                    break
            if not find:
                weekly_report.append((d, False))
        return weekly_report

    def create_draft(self, message_body, user_id='me'):
        message = {'message': message_body}
        draft = self.service.users().drafts().create(userId=user_id, body=message).execute()
        if draft:
            self.logger.debug('Draft has been created, please make a check.')
        return draft

def main():
    gmail = GamilAPI()
    #     # gmail.count_daily_with_email('yunyi@nonda.us')
    subject = settings.REMIND_SUBJECT
    message_text = settings.REMIND_TEXT
    message = gmail.create_message('xiaoxi@nonda.us', 'ganbinwen@nonda.me,binwengan@gmail.com', subject, message_text)
    gmail.send_message(message)

if __name__ == '__main__':
    main()






