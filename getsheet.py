#!/usr/bin/env python

#
# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
#

import google.oauth2.credentials

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow

# TODO: You will need to change this to the location of your secrets file
CLIENT_SECRETS_FILE = "/home/ross/research/oauth2_secrets/client_secrets.json"

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'

# The sheet-id is the identifier of the spreadsheet
#
# TODO: You will need to replace this with the ID of your spreadsheet
#       If can be found from the URL of the spreadsheet in your browser (between slashes)
SHEET_ID = '1DC8IB3duF9PFcYHklOLGNc40VQVLxGHiyYdy2CxTtmg'


class SpreadSheetRow(object):
    def __init__(self, apiRow):
        self.values = []
        for col in apiRow:
            self.values.append(col['formattedValue'])

    def htmlRow(self):
        result = "<tr>\n"
        for value in self.values:
            result += "  <td>%s</td>\n" % str(value)
        result += "</tr>\n"
        return result

    
class SpreadSheet(object):
    def __init__(self, apiRowData):
        self.rows = []
        for row in apiRowData:
            self.rows.append(SpreadSheetRow(row['values']))

    def htmlTable(self):
        result = "<table>\n"
        for row in self.rows:
            result += row.htmlRow()
        result += "</table>"
        return result


def get_authenticated_service():
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
    credentials = flow.run_local_server(host='localhost',
                                        port=8080, 
                                        success_message='You may close this window.',
                                        open_browser=True)
    return build(API_SERVICE_NAME, API_VERSION, credentials=credentials)


def get_sheet(service, **kwargs):
    results = service.spreadsheets().get(**kwargs).execute()
    return results


def extractCells(sheet):
    return SpreadSheet(sheet['sheets'][0]['data'][0]['rowData'])


if __name__ == '__main__':
    service     = get_authenticated_service()
    sheet       = get_sheet(service, spreadsheetId=SHEET_ID, includeGridData=True)
    spreadsheet = extractCells(sheet)
    htmlText    = spreadsheet.htmlTable()
    print(htmlText)
