
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
    credential_path = os.path.join(
        credential_dir,
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
    result = 1 if subprocess.check_output(
        ['./check-connection.sh']).split('\n')[0] == 'Online' else 0
    output = {
        'start': datetime.datetime.now(),
        'up': result,
        'down': 1 - result}
    return output


def runSpeedTest():
    speedtestOutput = subprocess.check_output(['/usr/local/bin/speedtest',
                                               '--simple']).split('\n')
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


def dateToEpoch(d):
    delta = d - datetime.datetime(1899, 12, 30, 0, 0)
    return delta.total_seconds() / (3600 * 24)


def iso8601stringToDate(d):
    """Parse ISO8601 datetime string (with optional milliseconds) to date
    """
    # our strptime format requires milliseconds, so add if not present
    if '.' not in d:
        d += '.000000'
    return datetime.datetime.strptime(d, '%Y-%m-%dT%H:%M:%S.%f')


def connectionTestRow(result):
    """Results are stored in form
    {start}, {end}, {success count}, {failure count}
    """
    return {
        'values': [
            {
                'userEnteredValue': {'numberValue': dateToEpoch(result['start'])}
            }, {
                'userEnteredValue': {'numberValue': dateToEpoch(result['end'])}
            }, {
                'userEnteredValue': {'numberValue': result['up']}
            }, {
                'userEnteredValue': {'numberValue': result['down']}
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
                'rowIndex': 1,
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
    output = []
    if os.path.exists(RESULT_QUEUE_FILE):
        with open(RESULT_QUEUE_FILE, 'r') as file:
            for line in file.read().splitlines():
                raw = json.loads(line)
                raw['start'] = iso8601stringToDate(raw['start'])
                if 'end' in raw:
                    raw['end'] = iso8601stringToDate(raw['end'])
                output += [raw]
    return output


def queueResult(thisResult):
    """Append test result to the queue if we could not upload this time
    """
    if thisResult:
        with open(RESULT_QUEUE_FILE, 'a') as file:
            file.write(json.dumps(thisResult, default=json_serial) + '\n')
        return True
    else:
        return False


def json_serial(obj):
    if isinstance(obj, datetime.datetime):
        serial = obj.isoformat()
        return serial
    raise TypeError('Type not serializable')


def clearResultQueue():
    """Clear queued test results after successful upload
    """
    with open(RESULT_QUEUE_FILE, 'w') as file:
        return True


def collapseResults(results):
    """Instead of one row per result, collapse by outcome, e.g. 30 consecutive
    successes = one row with 30 in success column. This will make results far
    more compact and readable
    """
    collapsedResults = []
    lastResult = None
    for result in results:
        if lastResult and ((result['up'] > 0 and lastResult['up'] > 0) or
                           (result['down'] > 0 and lastResult['down'] > 0)):
            lastResult = collapsedResults[len(collapsedResults) - 1]
            lastResult['up'] += result['up']
            lastResult['down'] += result['down']
        else:
            lastResult = result
            collapsedResults = collapsedResults + [result]
        lastResult['end'] = result['start']
    return collapsedResults


def sendResultsToGoogle(results):
    """Turn results (array of JSON test result objects) into Google Sheets
    API request and submit it.
    """
    # get last row number used
    lastRowNum = int(getRange('Sheet1!G2')[0][0])

    collapsedResults = collapseResults(results)

    # get last row from sheet
    lastRow = getRange('Sheet1!A{0}:D{0}'.format(lastRowNum))[0]
    # destructure it
    (lastStart, lastEnd, lastUp, lastDown) = lastRow
    # force blanks to zero
    lastUp = int(lastUp or 0)
    lastDown = int(lastDown or 0)

    # if same result as last time (up or down), merge saved values into
    # queued results and insert one row higher than planned to overwrite
    # that row in the sheet
    if (lastUp > 0 and collapsedResults[0]['up'] > 0) or \
            (lastDown > 0 and collapsedResults[0]['down'] > 0):
        collapsedResults[0]['start'] = datetime.datetime.strptime(lastStart, '%m/%d/%Y %H:%M:%S')
        collapsedResults[0]['up'] += lastUp
        collapsedResults[0]['down'] += lastDown
        lastRowNum = lastRowNum - 1  # rewind one row to overwrite

    # generate request array to insert result(s) one row after last
    requests = [setRows(lastRowNum + 1,
                        map(lambda result:
                            connectionTestRow(result), collapsedResults))]
    # increment last row number by total rows added
    requests.append(setLastRow(lastRowNum + len(collapsedResults)))

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
        print('{} - {}'.format(thisResult['start'],
                               'Up' if thisResult['up'] else 'Down'))
    except Exception as e:
        # Append test result (if any) to queue of pending results
        queueResult(thisResult)
        print('{} - ERROR: {}'.format(datetime.datetime.now(), e))
        traceback.print_exc(file=sys.stderr)

if __name__ == '__main__':
    main()
