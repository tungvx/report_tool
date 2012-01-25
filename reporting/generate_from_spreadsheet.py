try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import atom

gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client.email = 'toilatung90@gmail.com'
gd_client.password = 'tungyeungoc'
gd_client.source = 'exampleCo-exampleApp-1'
gd_client.ProgrammaticLogin()

def generate_from_spreadsheet(key):
    message = 'ok' #message to be returned to indicate whether the function is executed successfully
    feed = gd_client.GetCellsFeed(key, 1)
    PrintFeed(feed)
    gd_client.UpdateCell(row=2, col=2, inputValue="tung",key=key, wksht_id=1)
    
    return message #return the message  

def PrintFeed(feed):
  for i, entry in enumerate(feed.entry):
    if isinstance(feed, gdata.spreadsheet.SpreadsheetsCellsFeed):
      print '%s %s\n' % (entry.title.text, entry.content.text)
    elif isinstance(feed, gdata.spreadsheet.SpreadsheetsListFeed):
      print '%s %s %s' % (i, entry.title.text, entry.content.text)
      # Print this row's value for each column (the custom dictionary is
      # built from the gsx: elements in the entry.) See the description of
      # gsx elements in the protocol guide.
      print 'Contents:'
      for key in entry.custom:
        print '  %s: %s' % (key, entry.custom[key].text)
      print '\n',
    else:
      print '%s %s\n' % (i, entry.title.text)