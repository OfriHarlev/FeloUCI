import httplib2
import os
import time

from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_ID.json'
APPLICATION_NAME = 'Felo UCI'


class GoogleSheetsManager:
    def __init__(self):
        self._credentials = _get_credentials()
        http = self._credentials.authorize(httplib2.Http())
        discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        self._SHEETS = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discovery_url)

    def store_scores(self, group_name: str, scored_fencers: "[(Fencer, score),...,(Fencer, score)]"):
        """
        Creates a google sheet spreadsheet with the title of Group Name: {Current Time}
        and a list of scored fencers
        """
        init_sheet_data = {'properties': {'title': '{}: {}'.format(group_name, time.ctime())}}
        res = self._SHEETS.spreadsheets().create(body=init_sheet_data).execute()
        sheet_id = res['spreadsheetId']
        self._SHEETS.spreadsheets().values().update(spreadsheetId=sheet_id,
                                                    range="A1",
                                                    body={'values': [[group_name], [time.ctime()], [""],
                                                                     ["Name", "Score"]]},
                                                    valueInputOption='RAW').execute()
        self._SHEETS.spreadsheets().values().update(spreadsheetId=sheet_id,
                                                    range="A5",
                                                    body={'values': scored_fencers},
                                                    valueInputOption='RAW').execute()
        print("Created {}".format(res['properties']['title']))


def _get_credentials():
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
                                   'sheets.googleapis.com-python-FeloUCI.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


if __name__ == '__main__':
    GoogleSheetsManager().store_scores("Felo Fencing UCI", [("Matthew", 2000), ("Marco", 2000)])
