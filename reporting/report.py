import datetime
import django
from django.db import models
from django import forms
from xlwt.Workbook import Workbook
import xlrd,xlwt
import re
from report_tool.models import Pupil,Class,School
from xlutils.styles import Styles
from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils
from xlutils.filter import process,XLRDReader,XLWTWriter
import operator
import definitions
from itertools import groupby
import os

SITE_ROOT = os.path.dirname(os.path.realpath(__file__)) #path of the app
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded' #path to uploaded folder
FILE_GENERATE_PATH = SITE_ROOT + '/generated' #path to generated folder

#function to generate the report, receive the file name of the input file as the input
def generate(filename):
    fname = filename #name of the input file
    try:
        #extract the specified information
        function_name, index_of_function, head, index_of_head, body, indexes_of_body, input_file,index_of_excel_function, excel_function = fileExtractor(fname)
    except:
        return 'Wrong input file, please check all data' #if cannot extract the data, return wrong message
    else:
        try:
            #try to get the list of objects by executing the function in definitions.py file
            list_objects = eval('definitions.%s'%function_name)
        except :
            try:
                #execute the function directly
                list_objects = eval(function_name)
            except :
                #raise error if can not get list of objects
                return 'Definition of function error'
        try:
            #check if the function user specified returns the correct types of result
            len(list_objects)
        except :
            #if not, raise error
            return 'The function you defined returns wrong result (must return a list of objects)'

        #generate the report to the excel file, message here is the signal of the success
        message = generate_output(list_objects, index_of_function, head, index_of_head, body,
                                  indexes_of_body, input_file,fname, index_of_excel_function, excel_function)
        return message
    return 'ok'


#function to extract specifications from the template file
def fileExtractor(file):
    function_name = ''#name of the function which returns the list of objects
    head = '' #header
    index_of_head = [] #index of header
    index_of_function = [] #index of the function specification
    body = [] # contains the list of all the body data
    indexes_of_body = [] #indexes of the body data
    excel_function = [] #stores all the excel functions which user specified
    index_of_excel_function = [] #indexes of excel function
    fd = xlrd.open_workbook('%s/%s' % (FILE_UPLOAD_PATH, file), formatting_info=True)     #Read excel file for get data
    sheet = fd.sheet_by_index(0) # Get the first sheet
    
    #read information user specified
    for col_x in range(sheet.ncols):
        for row_x in range(sheet.nrows):
            value = sheet.cell(row_x,col_x).value # value in the excel file
            if value: #if the cell contains data
                temp = re.search('#<.*?>',unicode(value)) #if the cell contains the function which returns the data
                if temp:
                    function_name = (temp.group(0).rstrip('>').lstrip('#<')) #remove > at the right and #< at the left
                    index_of_function.append((row_x,col_x)) #stores the index of this function
                else:
                    temp = re.findall('{{.*?}}',unicode(value)) # find all the specified fields of data
                    if temp: #if yes
                        for temp1 in temp: #iterating all of the fields
                            temp1 = temp1.rstrip('}}').lstrip('{{') # remove tags to get attributes
                            if (temp1.startswith('head:')): #if the field is the header
                                head = temp1[5:] #remove head:
                                index_of_head.append((row_x,col_x)) #stores the location of the header
                            else:
                                body.append(temp1[5:]) #else the field is the body
                                indexes_of_body.append((row_x,col_x)) #stores the location of the body
                    if value.startswith(":="):
                        excel_function.append(value) #strores the value of the cell contain the specified excel function
                        index_of_excel_function.append((row_x,col_x)) #store index of above excel function
    return function_name, index_of_function, head, index_of_head, body, indexes_of_body,fd, index_of_excel_function, excel_function

def generate_output(list_objects,index_of_function,  head, index_of_head, body, indexes_of_body, input_file,fname, index_of_excel_function, excel_function):
    message = 'ok' #message to be returned to signal the success of the function

    #dict to store the values of the data fields. Dict here is used for grouping the data
    #the value of the header will be the keys of the dict
    dict = {}

    #manipulate the data
    message = manipulate_data(list_objects,index_of_function,  head, index_of_head, body, indexes_of_body, input_file,fname, index_of_excel_function, excel_function, dict)

    #if something's wrong, the return the message to raise exception
    if message != 'ok':
        return message
    
    keys =  sorted(dict.keys()) #sort the keys

    sheet = input_file.sheet_by_index(0) # Get the first sheet
    wtbook, style_list = copy2(input_file) #copy the content and the format(style) of the input file into wtbook
    wtsheet = wtbook.get_sheet(0)# get the first sheet of wtbook

    #remove the content at the position of the function which returns the data, remains the format of the cell
    xf_index = sheet.cell_xf_index(index_of_function[0][0],index_of_function[0][1])
    wtsheet.write(index_of_function[0][0],index_of_function[0][1],'', style_list[xf_index])

    #begin to write the data fields to wtbook
    row = 0 #variable used to travel all the rows in the wtsheet
    start_row = 0 #the row which is the starting for copying the data of the rows between the row of the head
        # and the row of the body
    
    if len(index_of_head) != 0: # if the header is not empty, we start at the row of the header
        row = index_of_head[0][0]
        start_row = row + 1
    else: #else we start with the row above the body part
        row = indexes_of_body[0][0]-2
        start_row = row + 1

    for l in range(len(dict)):#iterate all the elements of the dict
        key = keys[l] #get the key
        if len(index_of_head) != 0: #if the header is not empty
            row_index = index_of_head[0][0] #get index of row of the header
            col_index = index_of_head[0][1]-1 #get index of column of the cell at the left of the header
            xf_index = sheet.cell_xf_index(row_index, col_index) #get the styles of that cell
            #copy the value and the formats of that cell to the current row and the same index
            #this is the part of the grouping data. The header is repeated at each key
            wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])
            col_index += 1 # the index of the column of the header
            xf_index = sheet.cell_xf_index(row_index, col_index)# style of the cell containing the header
            wtsheet.write(row,col_index,key, style_list[xf_index])#write the value of the header to this cell

        #copy the information in rows between the row of the head and the row of the body
        for row_index in range(start_row, indexes_of_body[0][0], 1):
            row  += 1 # increase the current row by one
            for col_index in range(sheet.ncols): #iterate all the columns
                xf_index = sheet.cell_xf_index(row_index, col_index) #the format of the copied cell
                #copy the value and the format to the current cell
                wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])

        #write data fields to wtsheet
        values = dict.get(key) #get the list of the data fields of this key
        for i in range(len(values)): #iterate the list to get all the data fields
            row += 1 #increase the current row
            #set height of the current row equal to the row of the spcified body row
            wtsheet.row(row).height = sheet.rowinfo_map.get(indexes_of_body[0][0]).height 
            for h in range(len(indexes_of_body)):#iterate all the fields
                col_index = indexes_of_body[h][1] #get the index of the column of this field
                value = values[i][h] # the value of the current data
                xf_index = sheet.cell_xf_index(indexes_of_body[h][0],indexes_of_body[h][1]) #the format of the cell
                #if the index of the current data is the index of one specified excel function
                if indexes_of_body[h] in index_of_excel_function:
                    #replace the data in the excel function for later formula
                    excel_function[index_of_excel_function.index(indexes_of_body[h])] = excel_function[index_of_excel_function.index(indexes_of_body[h])].replace('{{body:' + body[h] + '}}',value)
                else:# else just write the value of the data field to the cell
                    wtsheet.write(row,col_index,value, style_list[xf_index])

            #write excel functions to the output file:
            for h in range(len(index_of_excel_function)):
                col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
                #the format of the cell
                xf_index = sheet.cell_xf_index(index_of_excel_function[h][0],index_of_excel_function[h][1])
                #get the excel function:
                temp_excel_function = excel_function[h]
                #remove := at the beginning
                temp_excel_function = temp_excel_function[2:]
                # process error for string in the input of the excel function:
                temp_excel_function = temp_excel_function.replace(unichr(8220),'"').replace(unichr(8221),'"')
                # try to excecute the excel function as a python function, and write the result to the ouput sheet
                try:
                    value_of_excel_function = eval(temp_excel_function)
                    wtsheet.write(row,col_index,value_of_excel_function,style_list[xf_index])
                except : #if can not execute as a python function, we will try to parse it as a excel formula
                    try:
                        wtsheet.write(row,col_index,xlwt.Formula(temp_excel_function),style_list[xf_index])
                    except : #if all the two above cases are failed, the raise syntax error
                        message =  'Error in excel formula definition (at cell (' + str(index_of_excel_function[h][0] + 1) + ', '
                        message = message + str(index_of_excel_function[h][1] + 1)
                        message = message + ')): Syntax error '
                        return message

        row += 2 #each group are separated by one row, for beauty
    row -= 1

    #copy the information provided by user at the end of the report to the end of the output file
    for row_index in range(indexes_of_body[0][0] + 1, sheet.nrows, 1):
        if (sheet.rowinfo_map.get(row_index)):
            wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height #copy the height
        for col_index in range(sheet.ncols):
            xf_index = sheet.cell_xf_index(row_index, col_index) #the format of the copied cell
            #copy the value and the format
            wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])
        row += 1

    #save output
    wtbook.save('%s/%s' % (FILE_GENERATE_PATH, fname))
    return message

# This function is used for manipulating the data:
def manipulate_data(list_objects,index_of_function,  head, index_of_head, body, indexes_of_body, input_file,fname, index_of_excel_function, excel_function, dict):
    message = 'ok'

    # compute values of the data fields and put them into the dict
    for i in list_objects:
        result = [] #store the all the values of the data fields of an object
        key = '' #init the key for this object. If header is empty, then all the objects will have the same
                # key (''), then the data will not be grouped
        if head != '': #if the head is not empty
            try:
                key = eval('i.%s' % head) #try compute the value of the header
            except: #if there is error, then raise exceptions
                message =  'Error in head definition (at cell (' + str(index_of_head[0][0] + 1) + ', '
                message = message + str(index_of_head[0][1] + 1)
                message = message + ')): Object has no attribute '
                message = message + head + '; or the function you defined returns wrong result (must return a list of objects)'
                return message #return the message to signal the failure of the function
        for y in body: #iterate all the fields in the body part of this object
            try:
                result.append(eval('i.%s' % y)) #try to evaluate the value of the field and add them into the result
            except: # if error, raise exception and return the message
                index = body.index(y)
                message =  'Error in body definition (at cell (' + str(indexes_of_body[index][0] + 1) + ', '
                message = message + str(indexes_of_body[index][1] + 1)
                message = message + ')): Object has no attribute '
                message = message + y + '; or the function you defined returns wrong result (must return a list of objects)'
                return message
        result = tuple(result)# convert to tupple: [] to ()
        if dict.get(key): # if the key allready exists, trivially append the result to this key
            dict[key].append(result)
        else: #else create a  new key, and append the result
            dict[key] = []
            dict[key].append(result)
    return message

#This function is used for coping the contents of a excel file to an other one
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list