import datetime
from django.db import models
from django import forms
from xlwt.Workbook import Workbook
import xlrd,xlwt
import re
from xlutils.styles import Styles
from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils
from xlutils.filter import process,XLRDReader,XLWTWriter
import operator
from itertools import groupby
import os
from extract_information import extract_information, get_list_of_object
from django.http import HttpResponse, HttpResponseRedirect

SITE_ROOT = os.path.dirname(os.path.realpath(__file__)) #path of the app
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded' #path to uploaded folder
FILE_GENERATE_PATH = SITE_ROOT + '/generated' #path to generated folder

#function to generate the report, receive the file name of the input file as the input
def generate(filename, request):
    fname = filename #name of the input file
    response = HttpResponse(mimetype='application/ms-excel')
    response['Content-Disposition'] = u'attachment; filename=%s' % fname

    #read input file, style list:
    input_book = xlrd.open_workbook('%s/%s' % (FILE_UPLOAD_PATH, filename), formatting_info=True)     #Read excel file for get data
    style_list = copy2(input_book) #copy the content and the format(style) of the input file into wtbook
    #create output file:
    wtbook = xlwt.Workbook(encoding='utf-8') #create new workbook

    for i in range(input_book.nsheets):
        sheet = input_book.sheet_by_index(i) # Get the first sheet

        try:
            #extract the specified information
            function_name, index_of_function, group, index_of_group, body, indexes_of_body, index_of_excel_function, excel_function, body_input, index_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, once, index_of_once, once_input, index_of_once_input, reserve_postions = fileExtractor(sheet)
        except:
            return 'Wrong input file, please check all data', response #if cannot extract the data, return wrong message
        else:
            message, list_objects = get_list_of_object(function_name,index_of_function, request)

            if message != 'ok':
                return message, response
            #generate the report to the excel file, message here is the signal of the success
            message = generate_output(list_objects, index_of_function, group, index_of_group, body,
                                      indexes_of_body, fname, index_of_excel_function, excel_function,
                                      body_input, index_of_body_input,
                                      head, index_of_head, head_input, index_of_head_input,
                                      foot, index_of_foot, foot_input, index_of_foot_input, request,
                                      once, index_of_once, once_input, index_of_once_input,
                                      sheet, style_list, wtbook, reserve_postions)


            if message != 'ok':
                return message, response
    wtbook.save(response)
    if request.session.get('is_spreadsheet'):
        wtbook.save('%s/%s' % (FILE_GENERATE_PATH, fname))
    return 'ok', response

#function to extract specifications from the template file
def fileExtractor(sheet):
    function_name = ''#name of the function which returns the list of objects
    group = {} #group
    index_of_group = {} #index of group
    index_of_function = [] #index of the function specification
    body = [] # contains the list of all the body data
    indexes_of_body = [] #indexes of the body data
    excel_function = [] #stores all the excel functions which user specified
    index_of_excel_function = [] #indexes of excel function
    body_input = [] #store input value of body
    indexes_of_body_input = [] #store index of body input
    head = {}#store header
    index_of_head = {} #store indexes of head,
    head_input = {} #store head input
    index_of_head_input = {} #store index of head input
    foot = {}
    index_of_foot = {}
    foot_input = {}
    index_of_foot_input = {}
    once = []
    index_of_once = []
    once_input = []
    index_of_once_input = []
    reserve_postions = []
    
    #read information user specified
    for col_x in range(sheet.ncols):
        for row_x in range(sheet.nrows):
            value = sheet.cell(row_x,col_x).value # value in the excel file
            if value: #if the cell contains data

                #call the function to extract information
                temp_function_name = extract_information(index_of_function, index_of_group, body, indexes_of_body,index_of_excel_function, excel_function, value, row_x, col_x,[],[], body_input, indexes_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, once, index_of_once, once_input, index_of_once_input, group, reserve_postions)

                #append the function_name and the group
                function_name += temp_function_name
    return function_name, index_of_function, group, index_of_group, body, indexes_of_body, index_of_excel_function, excel_function, body_input, indexes_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, once, index_of_once, once_input, index_of_once_input, reserve_postions

def generate_output(list_objects,index_of_function,  group, index_of_group, body, indexes_of_body,fname, index_of_excel_function, excel_function, body_input, index_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, request, once, index_of_once, once_input, index_of_once_input, sheet, style_list, wtbook, reserve_postions):
    #dict to store the values of the data fields. Dict here is used for grouping the data
    #the value of the group will be the keys of the dict
    dict = {}

    #manipulate the data
    message = manipulate_data(list_objects, group, index_of_group, body, indexes_of_body, dict, head, index_of_head, foot, index_of_foot, once, index_of_once, once_input, index_of_once_input, request, index_of_excel_function, excel_function, 0, sheet)

    #if something's wrong, the return the message to raise exception
    if message != 'ok':
        return message

    wtsheet = wtbook.add_sheet(sheet.name, cell_overwrite_ok=True)# create new sheet named as of sheet

    #copy column widths to output file
    for i in range(sheet.ncols):
        wtsheet.col(i).width = sheet.computed_column_width(i)

    #if function data is not specified:
    if len(index_of_function) == 0:
        #just copy the content of input file to ouput file:
        for row_index in range(sheet.nrows):
            if (sheet.rowinfo_map.get(row_index)):
                wtsheet.row(row_index).height = sheet.rowinfo_map.get(row_index).height #copy the height
            for col_index in range(sheet.ncols):
                write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_index, sheet.cell(row_index, col_index).value)
        return message

    #get row of body part
    if len(indexes_of_body) != 0:
        row_of_body = indexes_of_body[0][0]
    else:
        row_of_body = sheet.nrows - 1

    #copy information between beginning of input file and row of body part:
    for row_index in range(row_of_body):
        if (sheet.rowinfo_map.get(row_index)):
                wtsheet.row(row_index).height = sheet.rowinfo_map.get(row_index).height #copy the height
        for col_index in range(sheet.ncols):
            write_to_sheet(row_index,col_index, sheet, wtsheet, style_list, row_index, sheet.cell(row_index, col_index).value)

    #remove the content at the position of the function which returns the data, remains the format of the cell
    write_to_sheet(index_of_function[0][0],index_of_function[0][1],sheet, wtsheet, style_list, index_of_function[0][0], '')

    #begin to write the data fields to wtbook
    if len(indexes_of_body) > 0:
        row = indexes_of_body[0][0] #variable used to travel all the rows in the wtsheet

        #call this function to recursively write the groups to ouput sheet
        row, message = write_groups_to_excel(list_objects,index_of_function,  group, index_of_group, body, indexes_of_body,fname, index_of_excel_function, excel_function, body_input, index_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, request, once, index_of_once, once_input, index_of_once_input, sheet, style_list,wtsheet, dict , row, 0, reserve_postions)
        if message != 'ok':
            return message
        max_row = indexes_of_body[0][0];
        for i in reserve_postions:
            if max_row < i[0]:
                max_row = i[0]

        row += max_row - indexes_of_body[0][0]
        for row_index in range(max_row + 1, sheet.nrows, 1):
            row += 1
            if (sheet.rowinfo_map.get(row_index)):
                wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height #copy the height
            for col_index in range(sheet.ncols): #iterate all the columns
                if (row_index, col_index) not in reserve_postions:
                    write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, sheet.cell(row_index, col_index).value)

    #write once_input to output file
    for i in range(len(once_input)):
        row_index = index_of_once_input[i][0]
        col_index = index_of_once_input[i][1]
        write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_index, once_input[i])

    #write excel functions in the once part to the output file:
    for h in range(len(index_of_excel_function)):
        if index_of_excel_function[h] in index_of_once:
            col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
            row_index = index_of_excel_function[h][0] # get row index of the cell contain excel function
            #get the excel function:
            temp_excel_function = excel_function[h]
            #remove := at the beginning
            temp_excel_function = temp_excel_function[2:]
            # process error for string in the input of the excel function:
            temp_excel_function = temp_excel_function.replace(unichr(8220), '"').replace(unichr(8221), '"')
            # try to execute the excel function as a python function, and write the result to the ouput sheet
            try:
                value_of_excel_function = eval(temp_excel_function)
                write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_index
                               , value_of_excel_function)
            except: #if can not execute as a python function, we will try to parse it as a excel formula
                try:
                    write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_index
                                   , xlwt.Formula(temp_excel_function))
                except: #if all the two above cases are failed, the raise syntax error
                    message = 'Error in excel formula, python function definition (at cell (' + str(
                        index_of_excel_function[h][0] + 1) + ', '
                    message = message + str(index_of_excel_function[h][1] + 1)
                    message = message + ')): Syntax error '
                    return message


    return message
def write_groups_to_excel(list_objects,index_of_function,  group, index_of_group, body, indexes_of_body,fname, index_of_excel_function, excel_function, body_input, index_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, request, once, index_of_once, once_input, index_of_once_input, sheet, style_list,wtsheet, dict_values, row, key_index, reserve_postions):    
    message = 'ok' #message to be returned to signal the success of the function

    group_key, key_all = get_group_key_and_key_all(group, key_index)

    if group.get(key_all):#if the group exists
        col_index = index_of_group.get(key_all)[1] #get index of column of the group
        row_index = index_of_group.get(key_all)[0] #get index of column of the group
        write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_index, '')
        row = row - (indexes_of_body[0][0] - row_index) #start write from row of group
        start_row = row_index + 1
    else:#else start write from row of body
        start_row = indexes_of_body[0][0]
        row = row - 1

    #get head input of this group:
    current_head_input = head_input.get(key_all)
    #and get foot input of this group:
    current_foot_input = foot_input.get(key_all)

    keys =  dict_values.keys() #get the keys of the dict

    for l in range(len(dict_values)): #iterate all the groups
        if current_head_input:
            temp_current_head_input = current_head_input[:]
        if current_foot_input:
            temp_current_foot_input = current_foot_input[:]

        key = keys[l] #get the key

        #if the group exists:
        if index_of_group.get(key_all):
            row_index = index_of_group[key_all][0]    #get the row index of the current group

            #set row height:
            if (sheet.rowinfo_map.get(row_index)):
                wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height #copy the height

            #copy all data of the row containing the group:
            for col_index in range(sheet.ncols):
                if (row_index, col_index) not in reserve_postions:
                    write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, sheet.cell(row_index, col_index).value)

            col_index = index_of_group[key_all][1] #get index of column of the group
            #copy the value and the formats of that cell to the current row and the same index
            #this is the part of the grouping data. The group is repeated at each key
            write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, '')

        #copy the information in rows between the row of the group and the row of the body
        for row_index in range(start_row, indexes_of_body[0][0], 1):
            row  += 1 # increase the current row by one
            if (sheet.rowinfo_map.get(row_index)):
                wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height #copy the height
            for col_index in range(sheet.ncols): #iterate all the columns
                if (row_index, col_index) not in reserve_postions:
                    write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, sheet.cell(row_index, col_index).value)

        #write data fields to wtsheet
        values = dict_values.get(key) #get the list of the data fields of this key

        head_values = values[0]#values of header

        foot_values = values[1] #values of foot

        body_values = values[2] #values of body part

        #replace value head_values into head input
        temp_current_excel_function = excel_function[:]
        if index_of_head.get(key_all):
            for h in range(len(index_of_head.get(key_all))):
                value = head_values[h]
                if index_of_head[key_all][h] in index_of_excel_function:
                    #replace the data in the excel function for later formula
                    temp_current_excel_function[index_of_excel_function.index(index_of_head[key_all][h])] = temp_current_excel_function[
                                                                                        index_of_excel_function.index(
                                                                                            index_of_head[key_all][
                                                                                            h])].replace(
                        '{{head' + key_all + ':' + head[key_all][h] + '}}', unicode(value))

                else:# else just replace the value into the body input
                    temp_current_head_input[index_of_head_input[key_all].index(index_of_head[key_all][h])] = temp_current_head_input[index_of_head_input[key_all].index(index_of_head[key_all][h])].replace('{{head' + key_all + ':' + head[key_all][h] + '}}', unicode(value))
            
            #write head values to output file:
            for h in range(len(index_of_head_input[key_all])):
                col_index = index_of_head_input[key_all][h][1]
                row_index = index_of_head_input[key_all][h][0]
                write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row - (indexes_of_body[0][0] - row_index) + 1, temp_current_head_input[h])

            #write excel functions in the head part to the output file:
            for h in range(len(index_of_excel_function)):
                if index_of_excel_function[h] in index_of_head[key_all]:
                    col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
                    row_index = index_of_excel_function[h][0] # get row index of the cell contain excel function
                    #get the excel function:
                    temp_excel_function = temp_current_excel_function[h]
                    #remove := at the beginning
                    temp_excel_function = temp_excel_function[2:]
                    # process error for string in the input of the excel function:
                    temp_excel_function = temp_excel_function.replace(unichr(8220), '"').replace(unichr(8221), '"')
                    # try to execute the excel function as a python function, and write the result to the ouput sheet
                    try:
                        value_of_excel_function = eval(temp_excel_function)
                        write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row - (indexes_of_body[0][0] - row_index) + 1
                                                               , value_of_excel_function)
                    except: #if can not execute as a python function, we will try to parse it as a excel formula
                        try:
                            write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row - (indexes_of_body[0][0] - row_index) + 1
                                                               , xlwt.Formula(temp_excel_function))
                        except: #if all the two above cases are failed, the raise syntax error
                            message = 'Error in excel formula, python function definition (at cell (' + str(
                                index_of_excel_function[h][0] + 1) + ', '
                            message = message + str(index_of_excel_function[h][1] + 1)
                            message = message + ')): Syntax error '
                            return message
        
        #write body values to output file:
        if type(body_values) is dict:
            row, message = write_groups_to_excel(list_objects,index_of_function,  group, index_of_group, body, indexes_of_body,fname, index_of_excel_function, excel_function, body_input, index_of_body_input, head, index_of_head, head_input, index_of_head_input, foot, index_of_foot, foot_input, index_of_foot_input, request, once, index_of_once, once_input, index_of_once_input, sheet, style_list,wtsheet, body_values, row + 1, key_index + 1, reserve_postions)
            if message != 'ok':
                return row, message
        else:
            increase_row = 1
            for i in range(len(body_values)): #iterate the list to get all the data fields
                temp_current_excel_function = excel_function[:]
                temp_body_input = body_input[:]

                row += increase_row #increase the current row
                #set height of the current row equal to the row of the spcified body row
                wtsheet.row(row).height = sheet.rowinfo_map.get(indexes_of_body[0][0]).height
                for h in range(len(indexes_of_body)):#iterate all the fields
                    value = body_values[i][h] # the value of the current data
                    #if the index of the current data is the index of one specified excel function
                    if indexes_of_body[h] in index_of_excel_function:
                        #replace the data in the excel function for later formula
                        temp_current_excel_function[index_of_excel_function.index(indexes_of_body[h])] = temp_current_excel_function[index_of_excel_function.index(indexes_of_body[h])].replace('{{' + body[h] + '}}',unicode(value))

                    else:# else just replace the value into the body input
                        temp_body_input[index_of_body_input.index(indexes_of_body[h])] = temp_body_input[index_of_body_input.index(indexes_of_body[h])].replace('{{' + body[h] + '}}',unicode(value))


                #write body_input to the output file:
                for h in range(len(index_of_body_input)):
                    col_index = index_of_body_input[h][1] #get current column index of body
                    row_index = index_of_body_input[h][0] #get current row index of body
                    #write to output file
                    temp_increase_row = write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, ' '.join(temp_body_input[h].split()))
                    if temp_increase_row > increase_row:
                        increase_row = temp_increase_row

                #write excel functions to the output file:
                for h in range(len(index_of_excel_function)):
                    if index_of_excel_function[h] in indexes_of_body:
                        col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
                        row_index = index_of_excel_function[h][0] # get row index of the cell contain excel function
                        #get the excel function:
                        temp_excel_function = temp_current_excel_function[h]
                        #remove := at the beginning
                        temp_excel_function = temp_excel_function[2:]
                        # process error for string in the input of the excel function:
                        temp_excel_function = temp_excel_function.replace(unichr(8220),'"').replace(unichr(8221),'"')
                        # try to execute the excel function as a python function, and write the result to the ouput sheet
                        try:
                            value_of_excel_function = eval(temp_excel_function)
                            #if the value of the function is "remove_row", the delete the current data row
                            if (value_of_excel_function == "remove_row"):
                                for temp_index in range(len(indexes_of_body)):
                                    #clear data and get increase row
                                    temp_increase_row = write_to_sheet(row_index, indexes_of_body[temp_index][1], sheet, wtsheet, style_list, row, "")
                                    if temp_increase_row > increase_row:
                                        increase_row = temp_increase_row
                                row -= 1
                                break
                            else: #else output the value of the function to the input file
                                temp_increase_row = write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, value_of_excel_function)
                                if temp_increase_row > increase_row:
                                    increase_row = temp_increase_row
                        except : #if can not execute as a python function, we will try to parse it as a excel formula
                            try:
                                temp_increase_row = write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, xlwt.Formula(temp_excel_function))
                                if temp_increase_row > increase_row:
                                    increase_row = temp_increase_row
                            except : #if all the two above cases are failed, the raise syntax error
                                message =  'Error in excel formula definition (at cell (' + str(index_of_excel_function[h][0] + 1) + ', '
                                message = message + str(index_of_excel_function[h][1] + 1)
                                message = message + ')): Syntax error '
                                return message

                #copy format of other cell in the body row
                row_index = index_of_body_input[0][0]
                for col_index in range(sheet.ncols):
                    if (row_index, col_index) not in reserve_postions:
                        write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, '')

        if index_of_foot.get(key_all):
            #write foot values to output file:
            max_foot_row = row
            #insert foot values to the output file:
            #replace value foot_values into foot input
            temp_current_excel_function = excel_function[:]
            for f in range(len(index_of_foot.get(key_all))):
                value = foot_values[f]
                if index_of_foot[key_all][f] in index_of_excel_function:
                    #replace the data in the excel function for later formula
                    temp_current_excel_function[index_of_excel_function.index(index_of_foot[key_all][f])] = temp_current_excel_function[
                                                                                        index_of_excel_function.index(
                                                                                            index_of_foot[key_all][
                                                                                            f])].replace(
                        '{{foot' + key_all + ':' + foot[key_all][f] + '}}', unicode(value))

                else:# else just replace the value into the body input
                    try:
                        temp_current_foot_input[index_of_foot_input[key_all].index(index_of_foot[key_all][f])] = temp_current_foot_input[index_of_foot_input[key_all].index(index_of_foot[key_all][f])].replace('{{foot' + key_all + ':' + foot[key_all][f] + '}}', unicode(value))
                    except :
                        temp_current_foot_input[index_of_foot_input[key_all].index(index_of_foot[key_all][f])] = temp_current_foot_input[index_of_foot_input[key_all].index(index_of_foot[key_all][f])].replace('{{foot' + key_all + ':' + foot[key_all][f] + '}}', str(value).decode('utf-8'))

            #write foot values to output file:
            for f in range(len(index_of_foot_input.get(key_all))):
                col_index = index_of_foot_input.get(key_all)[f][1]
                row_index = index_of_foot_input.get(key_all)[f][0]
                row_of_foot = row + (row_index - indexes_of_body[0][0])
                if row_of_foot > max_foot_row:
                    max_foot_row = row_of_foot
                write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_of_foot, temp_current_foot_input[f])
            
            #copy the information provided by user at the end of the report to the end of the output file
            temp_row = row
            for row_index in range(indexes_of_body[0][0] + 1,indexes_of_body[0][0] + max_foot_row - row + 1, 1):
                temp_row += 1
                if (sheet.rowinfo_map.get(row_index)):
                    wtsheet.row(temp_row).height = sheet.rowinfo_map.get(row_index).height #copy the height
                for col_index in range(sheet.ncols):
                    #copy the value and the format
                    if (row_index, col_index) not in reserve_postions:
                        write_to_sheet(row_index,col_index,sheet, wtsheet, style_list, temp_row, sheet.cell(row_index,col_index).value)

            #write excel functions in the head part to the output file:
            for h in range(len(index_of_excel_function)):
                if index_of_excel_function[h] in index_of_foot.get(key_all):
                    col_index = index_of_excel_function[h][1] # get column index of the cell contain excel function
                    row_index = index_of_excel_function[h][0] # get row index of the cell contain excel function
                    #get the excel function:
                    temp_excel_function = temp_current_excel_function[h]
                    #remove := at the beginning
                    temp_excel_function = temp_excel_function[2:]
                    # process error for string in the input of the excel function:
                    temp_excel_function = temp_excel_function.replace(unichr(8220), '"').replace(unichr(8221), '"')
                    # try to execute the excel function as a python function, and write the result to the ouput sheet
                    row_of_foot = row + (row_index - indexes_of_body[0][0])
                    try:
                        value_of_excel_function = eval(temp_excel_function)
                        write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_of_foot
                                                               , value_of_excel_function)
                    except: #if can not execute as a python function, we will try to parse it as a excel formula
                        try:
                            write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row_of_foot
                                                               , xlwt.Formula(temp_excel_function))
                        except: #if all the two above cases are failed, the raise syntax error
                            message = 'Error in excel formula, python function definition (at cell (' + str(
                                index_of_excel_function[h][0] + 1) + ', '
                            message = message + str(index_of_excel_function[h][1] + 1)
                            message = message + ')): Syntax error '
                            return message

        if l < len(dict_values) - 1:
            row = max_foot_row
            row += 1


    # return 1, 'not ok'
    return row, message

# This function is used for manipulating the data:
def manipulate_data(list_objects, group, index_of_group, body, indexes_of_body, dict, head, index_of_head, foot, index_of_foot, once, index_of_once, once_input, index_of_once_input, request, index_of_excel_function, excel_function, key, sheet):
    message = 'ok'

    if key == 0:
    #compute values for once:
        if len(list_objects) > 0:
            a = list_objects[0]
            for o in range(len(once)):
                try:
                    value = eval('a["%s"]' %once[o])
                except :
                    try:
                        value = eval('a.%s'%once[o])
                    except :
                        try:
                            value = eval(once[o])
                        except :
                            value = ''
                if index_of_once[o] in index_of_excel_function:
                    #replace the data in the excel function for later formula
                    try:
                        excel_function[index_of_excel_function.index(index_of_once[o])] = excel_function[
                                                                                          index_of_excel_function.index(
                                                                                              index_of_once[
                                                                                              o])].replace(
                            '{{once:' + once[o] + '}}', unicode(value))
                    except :
                        excel_function[index_of_excel_function.index(index_of_once[o])] = excel_function[
                                                                                          index_of_excel_function.index(
                                                                                              index_of_once[
                                                                                              o])].replace(
                            '{{once:' + once[o] + '}}', str(value).decode('utf-8'))
                else:
                    try:
                        once_input[index_of_once_input.index(index_of_once[o])] = once_input[index_of_once_input.index(index_of_once[o])].replace('{{once:' + once[o] + '}}', unicode(value))
                    except :
                        once_input[index_of_once_input.index(index_of_once[o])] = once_input[index_of_once_input.index(index_of_once[o])].replace('{{once:' + once[o] + '}}', str(value).decode('utf-8'))
        else:
            for o in range(len(once)):
                value = ''
                once_input[index_of_once_input.index(index_of_once[o])] = once_input[index_of_once_input.index(index_of_once[o])].replace('{{once:' + once[o] + '}}', unicode(value))

    #get groups tags
    group_key, key_all = get_group_key_and_key_all(group, key)

    for i in list_objects:
        temp_key = '' #init the key for this object. If group is empty, then all the objects will have the same
        # key (''), then the data will not be grouped
        if group_key != '': #if the group is not empty
            try:
                temp_key = eval('i["%s"]' % group_key) #try compute the value of the group
            except: #if there is error, then raise exceptions
                try:
                    temp_key = eval('i.%s'%group_key)
                except :
                    try:
                        temp_key = eval(group_key)
                    except :
                        message = 'Error in group definition at sheet ' + sheet.name + ', cell ' + xlwt.Utils.rowcol_to_cell(index_of_group[key_all][0],index_of_group[key_all][1])
                        message = message + ': Object has no attribute '
                        message = message + str(
                            group_key) + '; or the function you defined returns wrong result (must return a list of objects)'
                        return message #return the message to signal the failure of the function
        if dict.get(temp_key):
            dict[temp_key][2].append(i)
        else:
            dict[temp_key] = []

            head_result = [] #store values for header of each group
            if head.get(key_all):
                for h in head.get(key_all):
                    try: #try evaluate head value
                        head_result.append(eval('i["%s"]' % h))#for raw sql
                    except :
                        try: #for django models
                            head_value = eval('i.%s'%h)
                            if head_value != None:
                                head_result.append(head_value) #if head result is not None
                            else:
                                head_result.append('')
                        except :
                            try:
                                head_result.append(eval(h))
                            except :
                                index = head.get(key_all).index(h)
                                message =  'Error in head definition at sheet ' + sheet.name + ', cell ' + xlwt.Utils.rowcol_to_cell(index_of_head.get(key_all)[index][0],index_of_head.get(key_all)[index][1])
                                message = message + ': Object has no attribute '
                                message = message + h + '; or the function you defined returns wrong result (must return a list of objects)'
                                return message
            head_result = tuple(head_result)
            dict[temp_key].append(head_result)

            #store the values for footer:
            foot_result = []
            if foot.get(key_all):
                for f in foot.get(key_all):
                    try:#try to evaluate foot value
                        foot_result.append(eval('i["%s"]' % f)) #for raw sql
                    except :
                        try: #for django models
                            foot_value = eval('i.%s'%f)
                            if (foot_value != None):
                                foot_result.append(foot_value) #if the foot value s not None
                            else:
                                foot_result.append('')
                        except:
                            try:
                                foot_result.append(eval(f))
                            except :
                                index = foot.get(key_all).index(f)
                                message =  'Error in foot definition at sheet ' + sheet.name + ', cell ' + xlwt.Utils.rowcol_to_cell(index_of_foot.get(key_all)[index][0],index_of_foot.get(key_all)[index][1])
                                message = message + ': Object has no attribute '
                                message = message + f + '; or the function you defined returns wrong result (must return a list of objects)'
                                return message

            foot_result = tuple(foot_result)
            dict[temp_key].append(foot_result)

            dict[temp_key].append([])
            dict[temp_key][2].append(i)

    keys = dict.keys()
    for k in keys:
        sub_list_objects = dict.get(k)[2][:]
        if key < len(group.keys()) - 1:
            dict[k][2] = {}
            message = manipulate_data(sub_list_objects, group, index_of_group, body, indexes_of_body, dict[k][2], head, index_of_head,
                foot, index_of_foot, once, index_of_once, once_input, index_of_once_input, request, index_of_excel_function,
                excel_function, key + 1, sheet)
            if message != "ok":
                return message
        else:
            dict[k][2] = []
            for i in sub_list_objects:
                result = []
                for y in body: #iterate all the fields in the body part of this object
                    try:
                        result.append(eval('i["%s"]' % y)) #try to evaluate the value of the field and add them into the result
                    except: # if error, raise exception and return the message
                        try:
                            body_value = eval('i.%s'%y)
                            if body_value != None:
                                result.append(body_value)
                            else:
                                result.append('')
                        except :
                            try:
                                result.append(eval(y))
                            except :
                                index = body.index(y)
                                message =  'Error in body definition at sheet ' + sheet.name + ', cell ' + xlwt.Utils.rowcol_to_cell(indexes_of_body[index][0],indexes_of_body[index][1])
                                message = message + ': Object has no attribute '
                                message = message + y + '; or the function you defined returns wrong result (must return a list of objects)'
                                return message
                result = tuple(result)# convert to tupple: [] to ()
                dict[k][2].append(result)


    return message

def get_group_key_and_key_all(group, key):
    #get groups tags
    try:
        key_all = sorted(group.keys())[key]
        group_key = group.get(key_all)
    except :
        key_all = ''
        group_key = ''
    return group_key, key_all

#This function is used for coping the contents of a excel file to an other one
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.style_list

def is_merged(position, sheet):
    for crange in sheet.merged_cells:
        if position[0] == crange[0] and position[1] == crange[2]:
            return True, crange
    return False, ()


#this function is used for writing values to wtsheet, prevent merged cells
def write_to_sheet(row_index, col_index, sheet, wtsheet, style_list, row, value):
    merged, merged_range = is_merged((row_index, col_index), sheet)
    xf_index = sheet.cell_xf_index(row_index, col_index) #the format of the copied cell
    #copy the value and the format to the current cell
    if merged:
        wtsheet.write_merge(row, row + merged_range[1] - merged_range[0] - 1, merged_range[2], merged_range[3] - 1,
                            value, style_list[xf_index])
        return merged_range[1] - merged_range[0]
    else:
        wtsheet.write(row, col_index, value, style_list[xf_index])
        return 1