# -*- coding:utf-8 -*-
__author__ = "ganbin"
import optparse
import Gmail
import Sheet

def main():
    options = optparse.OptionParser(
        usage='python run.py [[-s, --send], [-d, --draft], [-u, -update]]',
                                    description="Automatically:\n Create weekly report"
                                                "\nUpdate Google Sheet \nAnd use report to create a draft in Gmail."
    )
    options.add_option('-s', '--send',
                       action='callback', callback=send_email,
                       help='Send email to remind Manager make a score')

    options.add_option('-d', '--draft', action='callback',
                       callback=make_draft, help="Create one draft in Draft Label.")

    options.add_option('-u', '--update', action='callback',
                       callback=update_google_sheet, help='Update Google Sheet.')

    opts, args = options.parse_args()


def send_email(option, opt_str, value, parser):
    gmail = Gmail.GamilAPI()
    subject = '【提醒】新人周得分填写'
    message_text = 'Hi Managers，<br>请尽早和新人进行上周目标及结果的1-1沟通，并于今天5:00PM前更新新人上周的得分。<br>' \
                   '<div class="gmail_chip gmail_drive_chip" style="width:396px;height:18px;max-height:18px;background-color:rgb(245,245,245);padding:5px;color:rgb(34,34,34);font-family:arial;font-style:normal;font-weight:bold;font-size:13px;border:1px solid rgb(221,221,221);line-height:1">' \
                   '<a href="' + 'https://drive.google.com/drive/u/1/folders/0B2TOYKCbyuaRLWFnZUw2RVV3OXc' + '" style="display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-decoration:none;padding:1px 0px;border:none;width:100%" target="_blank"><img style="vertical-align:bottom;border:none" src="https://ssl.gstatic.com/docs/doclist/images/icon_11_spreadsheet_list.png" class="CToWUd">&nbsp;<span dir="ltr" style="color:rgb(17,85,204);text-decoration:none;vertical-align:bottom">' + '试用期目标' + '</span></a></div>'
    message = gmail.create_message('xiaoxi@nonda.us', 'ganbinwen@nonda.me', subject, message_text)
    gmail.send_message(message)


def make_draft(option, opt_str, value, parser):
    gmail = Gmail.GamilAPI()
    match_date = gmail.get_match_date()
    sheet = Sheet.SheetAPI(match_date)
    weekly_values, report_lack_list, score_lack_list = sheet.get_weekly_values_with_gmail(gmail)
    sheet.update_weekly_report(weekly_values)
    sheet_title = sheet.update_sheet_title(match_date)
    subject = "New nondar Daily Report & Weekly Score Update"
    sheet_title.extend(weekly_values)
    html = sheet.make_table_html(sheet_title)
    message_text = sheet.make_message_text(report_lack_list, score_lack_list, html)
    message = gmail.create_message('xiaoxi@nonda.us', 'ganbinwen@nonda.me', subject, message_text)
    gmail.create_draft(message)


def update_google_sheet(option, opt_str, value, parser):
    gmail = Gmail.GamilAPI()
    match_date = gmail.get_match_date()
    sheet = Sheet.SheetAPI(match_date)
    weekly_values, report_lack_list, score_lack_list = sheet.get_weekly_values_with_gmail(gmail)
    sheet.update_weekly_report(weekly_values)
    sheet.update_sheet_title(match_date)


if __name__ == '__main__':
    main()