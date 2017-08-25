# -*- coding:utf-8 -*-
__author__ = "ganbin"
import logging

# Global Settings
SENDER_EMAIL = 'xiaoxi@nonda.us'

# Remind Email Settings
REMIND_TO_EMAIL = ''  # This argument get from google Sheet.
REMIND_TEXT = 'Hi New nondars & Managers, <br>' \
              '请及时进行上周目标及结果的1-1沟通，并于今天5:00PM前完成1-1沟通的表格填写（表2）及新人上周的得分更新（表1）。<br>' \
              '<div class="gmail_chip gmail_drive_chip" style="width:396px;height:18px;max-height:18px;' \
              'background-color:rgb(245,245,245);padding:5px;color:rgb(34,34,34);' \
              'font-family:arial;font-style:normal;font-weight:bold;' \
              'font-size:13px;border:1px solid rgb(221,221,221);line-height:1">' \
              '<a href="' + 'https://drive.google.com/drive/u/1/folders/0B2TOYKCbyuaRLWFnZUw2RVV3OXc' + \
              '" style="display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' \
              'text-decoration:none;padding:1px 0px;border:none;width:100%" target="_blank">' \
              '<img style="vertical-align:bottom;border:none" ' \
              'src="https://ssl.gstatic.com/docs/doclist/images/icon_11_spreadsheet_list.png"' \
              ' class="CToWUd">&nbsp;<span dir="ltr" style="color:rgb(17,85,204);' \
              'text-decoration:none;vertical-align:bottom">' + '试用期目标' + '</span></a></div>'
REMIND_SUBJECT = '【提醒】新人周得分填写'


# Result Draft Settings
RESULT_EMAIL = 'nondacn@nonda.us'
RESULT_TEXT = "Hi all，<br>以下是上周新人的daily report和weekly score情况：<br><br>" \
               "规则：<br>- Daily report以大家每天抄送HR的邮件为准，未收到daily report标黄；" \
               "<br>- 【新人】Daily report未及时发送者，每遗漏一次请在nonda微信群内发50元红包；" \
               "<br>- Weekly score以各位试用期目标中每位的Mgr评分为准，未评分标黄；" \
               "<br>- 【Mgr】Weekly score未及时评定者，每遗漏一次请在nonda微信群内发100元红包。" \
               "<br>本周情况：<br>"
RESULT_SUBJECT = 'New nondar Daily Report & Weekly Score Update'
NEW_LACK = '<br>- 【新人】漏写Daily Report：%s：%s次'
MGR_LACK = '<br>- 【Mgr】漏打分：%s：%s次'
NO_NEW_LACK = '<br>- 【新人】没有人有遗漏'
MGR_LACK_RULE = '<br>- 【Mgr】漏打分：%s：%s次'
NO_MGR_LACK = '<br>- 【Mgr】没有人有遗漏'
NO_ONE_LACK = '<br>- 没有人有遗漏'

# Google Sheets Settings
DATA_RANGE_NAME = '数据源!A2:K'
RESULT_RANGE_NAME = '结果!A1:K'
SCORE_RANGE_NAME = 'A:G'
TARGET_RANGE_NAME = 'C7'
DATA_SHEET_ID = '1pI2PT9jKVgx13slYAvN85OquIaRTflMoZZbRFHsFZ_A'
RESULT_TITLE_1 = ["新人", "Mgr", "Mentor", "入职日期", "试用期目标"]
RESULT_TITLE_2 = ['%s\nDaily Report', '%s-%s\n周得分']
RESULT_SHEET_ID = '1pI2PT9jKVgx13slYAvN85OquIaRTflMoZZbRFHsFZ_A'

# Match Rules
RE_DATE_RULE = r'\d{0,4}/?(\d{1,2}/\d{1,2})'
LACK_MARK = ''
INTERVAL_DAY = 7
FINDED_MARK = '✓'
LABEL_NAME = 'Daily Report'


