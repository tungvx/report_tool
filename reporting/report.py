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
                temp = re.search('#<.*?>',unicode(value))
                if temp:
                    function_name = (temp.group(0).rstrip('>').lstrip('#<'))
                    index_of_function.append((row_x,col_x))
                else:
                    temp = re.findall('{{.*?}}',unicode(value))
                    if temp:
                        for temp1 in temp:
                            temp1 = temp1.rstrip('}}').lstrip('{{')
                            if (temp1.startswith('head:')):
                                head = temp1[5:]
                                index_of_head.append((row_x,col_x))
                            else:
                                body.append(temp1[5:])
                                indexes_of_body.append((row_x,col_x))
                    temp = re.search('#{.*?}',unicode(value))
                    if temp:
                        excel_function.append(value) #strores the value of the cell contain the specified excel function
                        index_of_excel_function.append((row_x,col_x)) #store index of above excel function
    return function_name, index_of_function, head, index_of_head, body, indexes_of_body,fd, index_of_excel_function, excel_function

def generate_output(list_objects,index_of_function,  head, index_of_head, body, indexes_of_body, input_file,fname, index_of_excel_function, excel_function):
    message = 'ok'
    sheet = input_file.sheet_by_index(0) # Get the first sheet
    wtbook, style_list = copy2(input_file)
    wtsheet = wtbook.get_sheet(0)
    xf_index = sheet.cell_xf_index(index_of_function[0][0],index_of_function[0][1])
    wtsheet.write(index_of_function[0][0],index_of_function[0][1],'', style_list[xf_index])
    dict = {}
    for i in list_objects:
        result = []
        key = ''
        if head != '':
            try:
                key = eval('i.%s' % head)
            except:
                message =  'Error in head definition (at cell (' + str(index_of_head[0][0] + 1) + ', '
                message = message + str(index_of_head[0][1] + 1)
                message = message + ')): Object has no attribute '
                message = message + head + '; or the function you defined returns wrong result (must return a list of objects)'
                return message
        for y in body:
            try:
                result.append(eval('i.%s' % y))
            except:
                index = body.index(y)
                message =  'Error in body definition (at cell (' + str(indexes_of_body[index][0] + 1) + ', '
                message = message + str(indexes_of_body[index][1] + 1)
                message = message + ')): Object has no attribute '
                message = message + y + '; or the function you defined returns wrong result (must return a list of objects)'
                return message
        result = tuple(result)
        if dict.get(key):
            dict[key].append(result)
        else:
            dict[key] = []
            dict[key].append(result)
    keys =  sorted(dict.keys())

    row = 0
    start_row = 0
    if len(index_of_head) != 0:
        row = index_of_head[0][0]
        start_row = row + 1
    else:
        row = indexes_of_body[0][0]-2
        start_row = row + 1

    for l in range(len(dict)):
        key = keys[l]
        if len(index_of_head) != 0:
            row_index = index_of_head[0][0]
            col_index = index_of_head[0][1]-1
            xf_index = sheet.cell_xf_index(row_index, col_index)
            wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])
            col_index += 1
            xf_index = sheet.cell_xf_index(row_index, col_index)
            wtsheet.write(row,col_index,key, style_list[xf_index])
        for row_index in range(start_row, indexes_of_body[0][0], 1):
            row  += 1
            for col_index in range(sheet.ncols):
                xf_index = sheet.cell_xf_index(row_index, col_index)
                wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])

        values = dict.get(key)
        for i in range(len(values)):
            row += 1
            wtsheet.row(row).height = sheet.rowinfo_map.get(indexes_of_body[0][0]).height
            for h in range(len(indexes_of_body)):
                col_index = indexes_of_body[h][1]
                value = values[i][h] # the value of the current data
                xf_index = sheet.cell_xf_index(indexes_of_body[h][0],indexes_of_body[h][1]) #the format of the cell
                #if the index of the current data is the index of one specified excel function
                if indexes_of_body[h] in index_of_excel_function:
                    #replace the data in the excel function for later formula
                    excel_function[index_of_excel_function.index(indexes_of_body[h])] = excel_function[index_of_excel_function.index(indexes_of_body[h])].replace('{{body:' + body[h] + '}}',value)
                else:
                    wtsheet.write(row,col_index,value, style_list[xf_index])

            #write excel functions to the output file:
            for h in range(len(index_of_excel_function)):
                col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
                #the format of the cell
                xf_index = sheet.cell_xf_index(index_of_excel_function[h][0],index_of_excel_function[h][1])
                #get the excel function:
                temp_excel_function = excel_function[h]
                #remove #{ at the beginning
                temp_excel_function = temp_excel_function[2:]
                #remove } at the end
                temp_excel_function = temp_excel_function[:len(temp_excel_function)-1]
                # process error for string in the input of the excel function:
                temp_excel_function = temp_excel_function.replace(unichr(8220),'"').replace(unichr(8221),'"')
                #write excel function
                wtsheet.write(row,col_index,xlwt.Formula(temp_excel_function),style_list[xf_index])

        row += 2
    row -= 1

    #set row-heights and column-widths
    for row_index in range(indexes_of_body[0][0] + 1, sheet.nrows, 1):
        if (sheet.rowinfo_map.get(row_index)):
            wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height
        for col_index in range(sheet.ncols):
            xf_index = sheet.cell_xf_index(row_index, col_index)
            wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])
        row += 1

    #save output
    wtbook.save('%s/%s' % (FILE_GENERATE_PATH, fname))
    return message

def generate(filename):
    fname = filename
    try:
        #extract the specified information
        function_name, index_of_function, head, index_of_head, body, indexes_of_body, input_file,index_of_excel_function, excel_function = fileExtractor(fname)
    except:
        return 'Wrong input file, please check all data'
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

        #generate the report to the excel file
        message = generate_output(list_objects, index_of_function, head, index_of_head, body,
                                  indexes_of_body, input_file,fname, index_of_excel_function, excel_function)
        return message
    return 'ok'

#This function is used for coping the contents of a excel file to an other one
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list