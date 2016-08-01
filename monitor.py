
from __future__ import print_function
import datetime
import httplib2
import json
import os
import subprocess
import sys
import traceback

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

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


def runConnectionTest():
    output = {
        'timestamp': str(datetime.datetime.now()),
        'connected': 1 if subprocess.check_output(
            ['./check-connection.sh']).split('\n')[0] == 'Online' else 0}
    return output


def runSpeedTest():
    speedtestOutput = subprocess.check_output(['/usr/local/bin/speedtest', '--simple']).split('\n')
    output = {'timestamp': str(datetime.datetime.now())}
    for field in speedtestOutput:
        # print('field [' + str(field) + ']')
        if(field.strip() != ''):
            (name, val) = field.split(':')
            output[name] = val.strip().split(' ')[0]
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


def connectionTestRow(connectionTestOutput):
    return {
        'values': [
            {
                'userEnteredValue': {'stringValue': connectionTestOutput['timestamp']}
            }, {
                'userEnteredValue': {'numberValue': connectionTestOutput['connected']}
            }
        ]
    }


def speedTestRow(timestamp, download, upload, ping):
    return {
        'values': [
            {
                'userEnteredValue': {'stringValue': timestamp}
            }, {
                'userEnteredValue': {'numberValue': download}
            }, {
                'userEnteredValue': {'numberValue': upload}
            }, {
                'userEnteredValue': {'numberValue': ping}
            }
        ]
    }


def setRows(rowNum, rows):
    return {
        'updateCells': {
            'start': {
                'sheetId': 0,
                'rowIndex': rowNum - 1,
                'columnIndex': 0
            },
            'rows': rows,
            'fields': 'userEnteredValue'
        }
    }


# def addRow(rowNum, speedtestResult):
#     return {
#         'updateCells': {
#             'start': {
#                 'sheetId': 0,
#                 'rowIndex': rowNum - 1,
#                 'columnIndex': 0
#             },
#             'rows': [
#                 {
#                     'values': [
#                         {
#                             'userEnteredValue': {'stringValue': str(datetime.datetime.now())}
#                         }, {
#                             'userEnteredValue': {'stringValue': speedtestResult['Download']}
#                         }, {
#                             'userEnteredValue': {'stringValue': speedtestResult['Upload']}
#                         }, {
#                             'userEnteredValue': {'stringValue': speedtestResult['Ping']}
#                         }
#                     ]
#                 }
#             ],
#             'fields': 'userEnteredValue'
#         }
#     }


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
    """Monitor internet connection, saving results to Google spreadsheet
    """
    # Cell to store last row used
    # Because the point of this is to test for connection failures,
    # we must expect saving to the Google spreadsheet will fail, so
    # test results must be queued to disk locally.
    thisResult = runConnectionTest()
    results = [thisResult]  # getLocallyQueuedResults().append(thisResult)
    try:
        rowNum = int(getRange('Sheet1!G1')[0][0]) + 1
        # print('A' * 60)
        # print(json.dumps())
        requests = [setRows(rowNum, map(lambda result: connectionTestRow(result), results))]
        requests.append(setLastRow(rowNum))

        # results.append(setLastRow(rowNum))
        # requests = setRows(rowNum, results)
        batchUpdateRequest = {'requests': requests}
        # print('-' * 60)
        # print(json.dumps(batchUpdateRequest))
        # print('-' * 60)
        getService().spreadsheets().batchUpdate(spreadsheetId=spreadsheetId,
                                                body=batchUpdateRequest).execute()
        # clearLocalResults()
        # rowNum = int(getRange('Sheet1!G1')[0][0]) + 1
        # speedtestResult = runSpeedTest(rowNum % 10 == 0)
        # saveResultsLocally(requests)
    except Exception as e:
        # saveResultsLocally(queuedResults)
        # print(json.dumps(requests))
        eprint('ERROR: ' + str(e))
        traceback.print_exc(file=sys.stderr)
        # DEAL WITH IT HERE
    # requests.append(addRow(rowNum, speedtestResult))
    # requests.append(setLastRow(rowNum))

if __name__ == '__main__':
    main()
