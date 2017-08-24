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
import gmail_sample


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
class SheetAPI(object):
    def __init__(self, match_date):
        self._scopes = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/spreadsheets',
        ]
        self._client_secret_file = 'client_sheets_secret.json'
        self._application_name = 'Google Sheets API Python Quickstart'
        self.service = self.get_service()
        self.match_date = match_date
        self.values = self._get_commer_values()

    def _get_credentials(self):
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-quickstart.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self._client_secret_file, self._scopes)
            flow.user_agent = self._application_name
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_service(self):
        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        service = discovery.build('sheets', 'v4', http=http,
                                  discoveryServiceUrl=discoveryUrl)
        return service

    def _get_commer_values(self):
        values_range_name = '数据源!A2:K'
        spreadsheetId = '1pI2PT9jKVgx13slYAvN85OquIaRTflMoZZbRFHsFZ_A'
        values = self._list_value_with_range_sheetId(values_range_name, spreadsheetId)
        return values

    def get_weekly_values_with_gmail(self, gmail):
        # 新人0 新人邮箱1 试用期目标地址2 Mgr名称3 Mgr邮箱地址4 mentor名称5 入职日期6 转正7 离职8 状态9
        values = self.values
        weekly_values = []
        target_address = 2
        new_name = 0
        new_email = 1
        mgr_name = 3
        mgr_email = 4
        mentor_name = 5
        entry_date = 6
        report_lack_list = []
        score_lack_list  = []
        for value in values:
            daily_value = []
            report_lack_times = 0
            score_lack_times = 0
            daily_value.append(value[new_name])
            daily_value.append(value[mgr_name])
            daily_value.append(value[mentor_name])
            daily_value.append(value[entry_date])
            if value[new_email]:
                daily_count = []
                weekly_count = gmail.count_daily_with_email(value[new_email])
                score_date = weekly_count[0][0]
                for d in weekly_count:
                    if d[1]:
                        daily_count.append('✓')
                    else:
                        daily_count.append('')
                        report_lack_times += 1
            else:
                continue
            re_sheetid = re.search('d/(.+)?/', value[target_address])

            if re_sheetid:
                sheetid = re_sheetid.group(1)
                target_range_name = 'C7'
                score_range_name = 'A:G'
                score = self.get_score(score_range_name, sheetid, score_date)
                if not score:
                    score_lack_times += 1
                target = self._list_value_with_range_sheetId(target_range_name, sheetid)
                if target:
                    target = target[0][0]
                else:
                    target = ''
                daily_value.append(target)
                daily_value.extend(daily_count)
                daily_value.append(score)
            report_lack_list.append((value[new_name], report_lack_times))
            score_lack_list.append((value[new_name], score_lack_times))
            weekly_values.append(daily_value)
        print(weekly_values)
        return weekly_values, report_lack_list, score_lack_list



    def _list_value_with_range_sheetId(self, range_name, spreadsheetId, majorDimension='ROWS'):
        values = []
        result = self.service.spreadsheets().values().get(spreadsheetId=spreadsheetId,
                                                     range=range_name,
                                                     majorDimension=majorDimension).execute()
        if result.get('values'):
            values = result['values']
        return values

    def update_sheet_title(self, match_date):
        sheet_title = ["新人", "Mgr", "Mentor", "入职日期", "试用期目标"]
        daily_report_title = []
        week_score_title = '%s-%s\n周得分' % (match_date[0], match_date[-1])
        for d in match_date:
            daily_report_date = '%s\nDaily Report' % d
            daily_report_title.append(daily_report_date)
        sheet_title.extend(daily_report_title)
        sheet_title.append(week_score_title)
        range_name = '工作表1!A1:K1'
        body = {
            'values': [sheet_title],
        }
        return self._write_weekly_values(range_name, body=body)

    def _write_weekly_values(self, range_name, body):
        # range_name = '工作表1!A2:K'
        spreadsheet_id = '1nPB6SMvRcXSD-4ELcKIqjp25G1csT7bcKXXvkGs64-A'
        result = self.service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body).execute()
        return result

    def update_weekly_report(self, weekly_values):
        range_name = '工作表1!A2:K'
        body = {
            'values': weekly_values
        }
        self._write_weekly_values(range_name, body)

    def get_score(self, range_name, spreadsheetId, date):
        values = self._list_value_with_range_sheetId(range_name, spreadsheetId, 'COLUMNS')
        score = ''
        print(values)
        for v in values:
            for d in v:
                if date in d:
                    index_date = v.index(d)
                    score = v[index_date + 1]
                    break
        return score



def make_table_html(weekly_values):
    result = '<table cellspacing="0" cellpadding="0" dir="ltr" border="1" style="table-layout:fixed;font-size:13px;font-family:arial,sans,sans-serif;border-collapse:collapse;border:none">'
    result += '<colgroup><col width="209"><col width="209"><col width="80"><col width="86"><col width="86"><col width="93"><col width="197"><col width="197"><col width="93"></colgroup>'
    result += '<tbody>'
    for value in weekly_values:
        result += '<tr style="height:21px">'
        for i in value:
            if i:
                result += '<td style="padding:2px 3px;background-color:rgb(207, 226, 243);border-color:rgb(0,0,0);font-family:arial;font-weight:bold;word-wrap:break-word;vertical-align:top;text-align:center" rowspan="1" colspan="2">' + i + '</td>'
            else:
                result += '<td style="padding:2px 3px;background-color:rgb(255, 252, 51);border-color:rgb(0,0,0);font-family:arial;font-weight:bold;word-wrap:break-word;vertical-align:top;text-align:center" rowspan="1" colspan="2">' + i + '</td>'
            result += '</td>'
        result += "</tr>"
    result += '</tbody>'
    result += "</table>"

    return result


def make_message_text(report_lack_list, score_lack_list, html):
    text = "Hi all，<br>以下是上周新人的daily report和weekly score情况：<br><br>" \
           "规则：<br>- Daily report以大家每天抄送HR的邮件为准，未收到daily report标黄；" \
           "<br>- 【新人】Daily report未及时发送者，每遗漏一次请在nonda微信群内发50元红包；" \
           "<br>- Weekly score以各位试用期目标中每位的Mgr评分为准，未评分标黄；" \
           "<br>- 【Mgr】Weekly score未及时评定者，每遗漏一次请在nonda微信群内发100元红包。" \
           "<br>本周情况：<br>"
    report_lacked = False
    score_lacked = False
    for r in report_lack_list:
        if r[1] == 0:
            continue
        else:
            report_lacked = True
            text += '<br>- 【新人】漏写Daily Report：%s：%s次' % (r)
    for s in score_lack_list:
        if s[1] == 0:
            continue
        else:
            score_lacked = True
            text += "<br>- 【Mgr】漏打分：%s：%s次" % (s)
    if not report_lacked and not score_lacked:
        text += '<br>- 没有人有遗漏'
    elif not report_lacked:
        text += '<br>- 【新人】没有人有遗漏'
    elif not score_lacked:
        text += '<br>- 【Mgr】没有人有遗漏'
    text += html
    return text


if __name__ == '__main__':
    gmail = gmail_sample.GamilAPI()
    match_date = gmail.get_match_date()
    sheet = SheetAPI(match_date)
    weekly_values, report_lack_list, score_lack_list = sheet.get_weekly_values_with_gmail(gmail)
    sheet.update_weekly_report(weekly_values)
    sheet_title = sheet.update_sheet_title(match_date)
    subject = "New nondar Daily Report & Weekly Score Update"


    title_html = make_table_html(sheet_title)
    body_html = make_table_html(weekly_values)
    html = title_html + body_html

    message_text = make_message_text(report_lack_list, score_lack_list, html)
    message = gmail.create_message('xiaoxi@nonda.us', 'ganbinwen@nonda.me', subject, message_text)
    gmail.send_message(message)




