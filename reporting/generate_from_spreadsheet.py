try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import atom
from extract_information import extract_information, get_list_of_object
from xlwt.Utils import cell_to_rowcol2

gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client.email = 'toilatung90@gmail.com'
gd_client.password = 'tungyeungoc'
gd_client.source = 'exampleCo-exampleApp-1'
gd_client.ProgrammaticLogin()

def generate_from_spreadsheet(key):
    message = 'ok' #message to be returned to indicate whether the function is executed successfully

    try: #try to get all the cell containing the data in the first sheet
        feed = gd_client.GetCellsFeed(key, 1)
    except :
        message = "wrong spreadsheet link, please check again"
        return message

    #extract information from the spreadsheet
    function_name, index_of_function, head, index_of_head, body, indexes_of_body,index_of_excel_function, excel_function = extract_file(feed)

    #get the list of objects containing the data
    message, list_of_objects = get_list_of_object(function_name,index_of_function)

    if message != 'ok': #if the operation is not successful
        return message #return the message

    #generate the report
    message = generate_output(list_of_objects, index_of_function, head, index_of_head, body,
                                  indexes_of_body, index_of_excel_function, excel_function, key)

    return message #return the message


#function to generate the report
def generate_output(list_of_objects, index_of_function, head, index_of_head, body, indexes_of_body, index_of_excel_function, excel_function, key):
    message = 'ok' #message the indicate the success of the function

    
    
    return message #return the message
#        print '%s %s\n' % (entry.title.text, entry.content.text)
#    gd_client.UpdateCell(row=2, col=2, inputValue="tung",key=key, wksht_id=1)


#function to extract information from the spreadsheet
def extract_file(feed):
    function_name = ''#name of the function which returns the list of objects
    head = '' #header
    index_of_head = [] #index of header
    index_of_function = [] #index of the function specification
    body = [] # contains the list of all the body data
    indexes_of_body = [] #indexes of the body data
    excel_function = [] #stores all the excel functions which user specified
    index_of_excel_function = [] #indexes of excel function

    for entry in feed.entry: #iterate all the cells and extract information from each cell
        #get the value, row, column of the cell
        value = entry.content.text
        value = value.decode('utf-8')

        #convert the cell letter to the cell number
        temp_position = cell_to_rowcol2(entry.title.text)

        #get the row number and column number of the cell
        row_x = temp_position[0]
        col_x = temp_position[1]
        
        #call the function to extract information
        temp_function_name, temp_head = extract_information(index_of_function, index_of_head, body, indexes_of_body,index_of_excel_function, excel_function, value, row_x, col_x)

        #append the function_name and the header
        function_name += temp_function_name
        head += temp_head

    #return all the computed values
    return function_name, index_of_function, head, index_of_head, body, indexes_of_body, index_of_excel_function, excel_function

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