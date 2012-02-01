try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree

import gdata
import gdata.spreadsheet.service
import gdata.service
import gdata.spreadsheet
import gdata.docs
import gdata.docs.data
import gdata.docs.client
import gdata.docs.service
import gdata.spreadsheet.service
import datetime
import os
from report import generate

SITE_ROOT = os.path.dirname(os.path.realpath(__file__)) #path of the app
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded' #path to uploaded folder
FILE_GENERATE_PATH = SITE_ROOT + '/generated' #path to generated folder

def generate_from_spreadsheet(key, token):
    message = 'ok' #message to be returned to indicate whether the function is executed successfully

    try: #try to get all the cell containing the data in the first sheet
        gd_client = gdata.docs.service.DocsService()
        gd_client.email = 'toilatung90@gmail.com'
        gd_client.password = 'password'
        gd_client.ssl = True
        gd_client.source = "My Fancy Spreadsheet Downloader"
        gd_client.ProgrammaticLogin()
        spreadsheets_client = gdata.spreadsheet.service.SpreadsheetsService()
        spreadsheets_client.email = gd_client.email
        spreadsheets_client.password = gd_client.password
        spreadsheets_client.source = "My Fancy Spreadsheet Downloader"
        spreadsheets_client.ProgrammaticLogin()
        uri = 'http://docs.google.com/feeds/documents/private/full/%s' % key
        entry = gd_client.GetDocumentListEntry(uri)
        title = entry.title.text
        docs_auth_token = gd_client.GetClientLoginToken()
        gd_client.SetClientLoginToken(spreadsheets_client.GetClientLoginToken())
        now = datetime.datetime.now()
        uploaded_file_name = str(now.year)+str(now.day)+str(now.month)+str(now.hour)+str(now.minute)+str(now.second) + '.xls'
        gd_client.Export(entry, FILE_UPLOAD_PATH + '/' + uploaded_file_name)
        gd_client.SetClientLoginToken(docs_auth_token)
    except :
        return "wrong spreadsheet link, please check again", "", ""

    #call generate function
    message = generate(uploaded_file_name)

    if  message != 'ok':
        return message, "", ""

    message, output_link = upload_result(uploaded_file_name, title)

    return message, output_link, title #return the message

def upload_result(file_name, title):
    message = 'ok'
    gd_client = gdata.docs.service.DocsService(source='yourCo-yourAppName-v1')
    gd_client.ClientLogin('toilatung90@gmail.com', 'password')
    ms = gdata.MediaSource(file_path=FILE_GENERATE_PATH + '/' + file_name, content_type=gdata.docs.service.SUPPORTED_FILETYPES['XLS'])
    entry = gd_client.Upload(ms, 'Report result of ' + title)
    output_link = entry.GetAlternateLink().href
    return message, output_link