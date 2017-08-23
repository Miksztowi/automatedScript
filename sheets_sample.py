from __future__ import print_function
# -*- coding:utf-8 -*-
__author__ = "ganbin"
import httplib2
import os
import re
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets',
]

CLIENT_SECRET_FILE = 'client_sheets_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    weekly_values = []
    # 新人0 新人邮箱1 试用期目标地址2 Mgr名称3 Mgr邮箱地址4 mentor名称5 入职日期6 转正7 离职8 状态9
    new_name = 0
    new_email = 1
    target_address = 2
    mgr_name = 3
    mgr_email = 4
    mentor_name = 5
    entry_date = 6
    values_range_name = 'A2:K'
    spreadsheetId = '1pI2PT9jKVgx13slYAvN85OquIaRTflMoZZbRFHsFZ_A'
    values = list_value_with_range_sheetId(service, values_range_name, spreadsheetId)
    for value in values:
        re_sheetid = re.search('d/(.+)?/', value[target_address])
        if re_sheetid:
            sheetid = re_sheetid.group(1)
            target_range_name = 'C7'
            target = list_value_with_range_sheetId(service, target_range_name, sheetid)
            if target:
                target = target[0][0]
            else:
                target = ''
            weekly_values.append([value[new_name],
                                  value[mgr_name],
                                  value[mentor_name],
                                  value[entry_date],
                                  target,
                                  ])
    write_weekly_values(service, weekly_values)


def list_value_with_range_sheetId(service, range_name, spreadsheetId):
    values = []
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId,
                                                 range=range_name,
                                                 majorDimension='ROWS').execute()
    if result.get('values'):
        values = result['values']
    return values


def write_weekly_values(service, weekly_values):
    body = {
        'values': weekly_values
    }
    range_name = '工作表2!A2:E'
    spreadsheet_id = '1nPB6SMvRcXSD-4ELcKIqjp25G1csT7bcKXXvkGs64-A'
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body).execute()
    print(result)


if __name__ == '__main__':
    main()