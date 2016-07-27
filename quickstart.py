
from __future__ import print_function
import datetime
import httplib2
import os
import subprocess

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

spreadsheetId = '1RU0hbJSlBBs_svv0XX__gt_gAuNveOs_s3a0DQHuVsM'

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
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

    store = oauth2client.file.Storage(credential_path)
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


def runSpeedTest():
    speedtestOutput = subprocess.check_output(['/usr/local/bin/speedtest', '--simple']).split('\n')
    # speedtestOutput = 'Ping: 37.37 ms\nDownload: 32.54 Mbit/s\nUpload: 5.67 Mbit/s\n'.split('\n')
    output = {}
    for field in speedtestOutput:
        # print('field [' + str(field) + ']')
        if(field.strip() != ''):
            (name, val) = field.split(':')
            output[name] = val.strip()
    return output


def getService():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    return service


def getRange(rangeName):
    service = getService()

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    return result.get('values', [])


def now_epoch_days():
    # epoch = datetime.datetime.utcfromtimestamp(0).total_seconds()
    epoch = 123
    return str(epoch)


def addRow(rowNum, speedtestResult):
    return {
        'updateCells': {
            'start': {
                'sheetId': 0,
                'rowIndex': rowNum - 1,
                'columnIndex': 0
            },
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredValue': {'stringValue': str(datetime.datetime.now())}
                        }, {
                            'userEnteredValue': {'stringValue': speedtestResult['Download']}
                        }, {
                            'userEnteredValue': {'stringValue': speedtestResult['Upload']}
                        }, {
                            'userEnteredValue': {'stringValue': speedtestResult['Ping']}
                        }
                    ]
                }
            ],
            'fields': 'userEnteredValue'
        }
    }


def setLastRow(rowNum):
    return {
        'updateCells': {
            'start': {
                'sheetId': 0,
                'rowIndex': 0,
                'columnIndex': 6
            },
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredValue': {'numberValue': rowNum}
                        }
                    ]
                }
            ],
            'fields': 'userEnteredValue'
        }
    }


def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    # values = getRange('Sheet1!A2:E')
    rowNum = int(getRange('Sheet1!G1')[0][0]) + 1
    # print('Last row: ' + lastRow[0][0])
    # rangeName = 'Sheet1!A' + lastRow + ':E' + lastRow
    # print('Range name: ' + rangeName)
    requests = []
    speedtestResult = runSpeedTest()
    requests.append(addRow(rowNum, speedtestResult))
    requests.append(setLastRow(rowNum))
    batchUpdateRequest = {'requests': requests}
    getService().spreadsheets().batchUpdate(spreadsheetId=spreadsheetId,
                                            body=batchUpdateRequest).execute()

#    values = getRange(rangeName)
#
#    if not values:
#        print('No data found.')
#    else:
#        print('Last row of values')
#        for row in values:
#            for col in row:
#                print('%s ' % (col))
#
#    print(createUpdateItem('potato', 'abc'))

if __name__ == '__main__':
    main()
