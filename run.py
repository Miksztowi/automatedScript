# -*- coding:utf-8 -*-
__author__ = "ganbin"
import optparse
import settings
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
    sheet = Sheet.SheetAPI()
    values = sheet.values
    subject = settings.REMIND_SUBJECT
    sender = settings.SENDER_EMAIL
    message_text = settings.REMIND_TEXT
    for value in values:
        if value[1] and value[4]:
            to = ','.join([value[1], value[4]])
            message = gmail.create_message(sender, 'ganbinwen@nonda.me,binwengan@gmail.com', subject, message_text)
            gmail.send_message(message)


def make_draft(option, opt_str, value, parser):
    gmail = Gmail.GamilAPI()
    match_date = gmail.get_match_date()
    sheet = Sheet.SheetAPI()
    weekly_values, report_lack_list, score_lack_list = sheet.get_weekly_values_with_gmail(gmail)
    sheet.update_weekly_report(weekly_values)
    sheet_title = sheet.update_sheet_title(match_date)
    sheet_title.extend(weekly_values)
    html = sheet.make_table_html(sheet_title)
    message_text = sheet.make_message_text(report_lack_list, score_lack_list, html)
    subject = settings.RESULT_SUBJECT
    sender = settings.SENDER_EMAIL
    to = settings.RESULT_EMAIL
    message = gmail.create_message(sender, to, subject, message_text)
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