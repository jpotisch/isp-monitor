
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
RESULT_QUEUE_FILE = 'resultQueue.txt'


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


def connectionTestRow(result):
    """Results are stored in form {timestamp}, {success}, {failure}
    """
    return {
        'values': [
            {
                'userEnteredValue': {'stringValue': result['timestamp']}
            }, {
                'userEnteredValue': {'numberValue': 1}
            } if result['connected'] == 1 else {}, {
                'userEnteredValue': {'numberValue': 1}
            } if result['connected'] == 0 else {}
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
    """Create Google Sheets API payload to set a range of cells
    """
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


def setLastRow(rowNum):
    """Store the last row number used in its own cell so we know where to
    insert our row(s) next time
    """
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


def getQueuedResults():
    """Returns array of queued test result JSON objects
    """
    if os.path.exists(RESULT_QUEUE_FILE):
        with open(RESULT_QUEUE_FILE, 'r') as file:
            return map(lambda line: json.loads(line), file.read().splitlines())
    else:
        return []


def queueResult(thisResult):
    """Append test result to the queue if we could not upload this time
    """
    if thisResult:
        with open(RESULT_QUEUE_FILE, 'a') as file:
            file.write(json.dumps(thisResult) + '\n')
        return True
    else:
        return False


def clearResultQueue():
    """Clear queued test results after successful upload
    """
    with open(RESULT_QUEUE_FILE, 'w') as file:
        return True


def sendResultsToGoogle(results):
    """Turn results (array of JSON test result objects) into Google Sheets
    API request and submit it.
    """
    # get last row number used
    lastRow = int(getRange('Sheet1!G1')[0][0])

    # generate request array to insert result(s) one row after last
    requests = [setRows(lastRow + 1, map(lambda result: connectionTestRow(result), results))]
    # increment last row number by total rows added
    requests.append(setLastRow(lastRow + len(results)))

    # submit request
    batchUpdateRequest = {'requests': requests}
    getService().spreadsheets().batchUpdate(spreadsheetId=spreadsheetId,
                                            body=batchUpdateRequest).execute()
    return True


def main():
    """Monitor internet connection, saving results to Google spreadsheet
    """
    # Because the point of this is to test for connection failures,
    # we must expect saving to the Google spreadsheet will fail, so
    # in case of failure, test results are queued for next time
    thisResult = None
    try:
        thisResult = runConnectionTest()

        # Retrieve any results queued from previous network failures
        allResults = getQueuedResults() + [thisResult]

        # Attempt to save to Google spreadsheet
        sendResultsToGoogle(allResults)

        # If we made it this far it worked, so clear the queue
        clearResultQueue()

        # Output result
        print('{} - {}'.format(thisResult['timestamp'], 'Up' if thisResult['connected'] else 'Down'))
    except Exception as e:
        # Append test result (if any) to queue of pending results
        queueResult(thisResult)
        print('{} - ERROR: {}'.format(datetime.datetime.now(), e))
        traceback.print_exc(file=sys.stderr)

if __name__ == '__main__':
    main()
