import arrow
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from googleapiclient import sample_tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secrets.json'
APPLICATION_NAME = 'newscrypto rewrite ranking tool'
DISCOVERY_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'


# SPREADHSEET SPECIFIC INFO
SHEET_ID = '1wE2YkNIT7wYuR7Ul23Xv98WqIQwe0gNnOuw3a-3jJpY'
TAB_NAME = 'リライトシート'
KEYWORD_RANGE = '%s!I2:I' % TAB_NAME
COLUMN_RANGE = '%s!1:1' % TAB_NAME

# COMMON CONSTANTS
URL = 'https://www.newscrypto.jp'
END_DATE = arrow.utcnow()
START_DATE = END_DATE.replace(days=-7)
FORMATTED_END_DATE = END_DATE.to('Asia/Tokyo').format('YYYY-MM-DD')
FORMATTED_START_DATE = START_DATE.to('Asia/Tokyo').format('YYYY-MM-DD')


def get_ranks(keywords):
    service, flags = sample_tools.init(
        [], 'webmasters', 'v3', __doc__, __file__,
        scope='https://www.googleapis.com/auth/webmasters')

    ranks = []
    for keyword in keywords:
        request = {
            'startDate': FORMATTED_START_DATE,
            'endDate': FORMATTED_END_DATE,
            'searchType': 'web',
            'dimensions': ['query'],
            'dimensionFilterGroups': [{
                'filters': [{
                    'dimension': 'country',
                    'expression': 'jpn'
                }, {
                    'dimension': 'query',
                    'operator': 'equals',
                    'expression': keyword
                }]
            }],
        }
        response = execute_request(service, URL, request)
        ranks.append(response.get('rows', [{}])[0].get('position', ''))

    return ranks


def execute_request(service, property_uri, request):
    return service.searchanalytics().query(
        siteUrl=property_uri, body=request
    ).execute()


def get_credentials():
    '''Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    '''
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(
        credential_dir, 'sheets.googleapis.com-python-quickstart.json'
    )

    store = Storage(credential_path)
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


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=DISCOVERY_URL)

    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=COLUMN_RANGE).execute()

    new_column_alpha = colnum_string(len(result.get('values', [[]])[0]) + 1)
    append_range = '%s!%s1' % (TAB_NAME, new_column_alpha)

    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=KEYWORD_RANGE).execute()
    keyword_values = result.get('values', [])

    keywords = [
        (k[0] if k else '').strip().lower().replace(
            '\u3000', ' '
        ) for k in keyword_values
    ]
    appending_data = [FORMATTED_END_DATE] + get_ranks(keywords)

    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID, range=append_range, body={
            'majorDimension': 'COLUMNS',
            'values': [appending_data]
        }, valueInputOption='RAW'
    ).execute()


if __name__ == '__main__':
    main()
