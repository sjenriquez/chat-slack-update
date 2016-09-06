import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import requests
import time

# Authorize gspread
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('</path/google-creds.json>', scope)
GC = gspread.authorize(credentials)
WKS = GC.open("<spreadsheet-name>").sheet1

def tomorrow_date():
	# Get tomorrow's date matching spreadsheet format, ex. Monday, Jul 11
	today = datetime.date.today()
	if today.strftime("%A") == "Friday":
		tomorrow = datetime.date.today() + datetime.timedelta(days=3)
	else:
		tomorrow = datetime.date.today() + datetime.timedelta(days=1)

	if tomorrow.strftime("%d")[0] == '0':
		day_num = tomorrow.strftime("%d")[1:]
		tomorrow_readable_format = tomorrow.strftime("%A, " + "%b " + day_num)
	else:
		tomorrow_readable_format = tomorrow.strftime("%A, %b %d")
	return tomorrow_readable_format

def get_tomorrow_cell(tomorrow_readable_format):
	# Find cell matching tomorrow's date
	tomorrow_cell = WKS.find(tomorrow_readable_format)
	return tomorrow_cell

def get_larkers(tomorrow_cell):
	# Get scheduled SE's, expects names in same column as date in the 4 cells directly beneath the date
	tomorrow_row_number = tomorrow_cell.row
	tomorrow_column_number = tomorrow_cell.col
	larkers = {}
	larkers['am1'] = WKS.cell(tomorrow_row_number+1, tomorrow_column_number).value
	larkers['am2'] = WKS.cell(tomorrow_row_number+2, tomorrow_column_number).value
	larkers['pm1'] = WKS.cell(tomorrow_row_number+3, tomorrow_column_number).value
	larkers['pm2'] = WKS.cell(tomorrow_row_number+4, tomorrow_column_number).value
	return larkers

def get_hacker(tomorrow_cell):
	# Expects "Hack day" row in one of 5 rows above the date row
	hacker = ""
	hack_week = False
	for i in xrange(1, 5):
		cell_value = WKS.cell(tomorrow_cell.row - i, 1).value
		if cell_value.lower() == "hack day":
			if WKS.cell(tomorrow_cell.row - i, tomorrow_cell.col).value != "":
				hacker = WKS.cell(tomorrow_cell.row - i, tomorrow_cell.col).value
				hack_week = True
			break
	return hacker, hack_week


def update_slack(hack_week, hacker, larkers):
	token='<slack_token>'
	if hack_week:
		support_topic = "AM 10-2:30 EDT: {0} + {1} | PM 2:30-7 EDT: {2} + {3} | Hack day: {4} | Olark Schedule: <spreadsheet-url>".format(larkers['am1'],larkers['am2'],larkers['pm1'],larkers['pm2'],hacker)
	else:
		support_topic = "AM 10-2:30 EDT: {0} + {1} | PM 2:30-7 EDT: {2} + {3} | Olark Schedule: <spreadsheet-url>".format(larkers['am1'],larkers['am2'],larkers['pm1'],larkers['pm2'])

	irc_topic = "AM 10-2:30 EDT: {0} + {1} | PM 2:30-7 EDT: {2} + {3} | #irc is a read-only mirror of #datadog on freenode".format(larkers['am1'],larkers['am2'],larkers['pm1'],larkers['pm2'])
	
	# "Make #support-team group call with topic
	support_team_update = requests.post('https://slack.com/api/groups.setTopic', data = {'token':token,'channel':'<slack-channel-id>','topic':support_topic})
	
	# Make #irc channel call with topic
	irc_update = requests.post('https://slack.com/api/channels.setTopic', data = {'token':token,'channel':'<slack-channel-id>','topic':irc_topic})
	
	# Make #scott-test channel call with topic
	scott_test_update = requests.post('https://slack.com/api/channels.setTopic', data = {'token':token,'channel':'<slack-channel-id>','topic':support_topic})

tomorrow_readable_format = tomorrow_date()
tomorrow_cell = get_tomorrow_cell(tomorrow_readable_format)
larkers = get_larkers(tomorrow_cell)
hacker, hack_week = get_hacker(tomorrow_cell)
update_slack(hack_week,hacker,larkers)