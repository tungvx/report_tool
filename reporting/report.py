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

SITE_ROOT = os.path.dirname(os.path.realpath(__file__))
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded'
FILE_GENERATE_PATH = SITE_ROOT + '/generated'

def fileExtractor(file):
    function_name = ''
    head = ''
    index_of_head = []
    index_of_function = []
    body = []
    indexes_of_body = []
    fd = xlrd.open_workbook('%s/%s' % (FILE_UPLOAD_PATH, file), formatting_info=True)     #Read excel file for get data
    sheet = fd.sheet_by_index(0) # Get the first sheet
    for col_x in range(sheet.ncols):
        for row_x in range(sheet.nrows):
            value = sheet.cell(row_x,col_x).value
            if value:
                temp = re.search('#<.*?>',unicode(value))
                if temp:
                    function_name = (temp.group(0).rstrip('>').lstrip('#<'))
                    index_of_function.append((row_x,col_x))
                    break
                temp = re.search('{{.*?}}',unicode(value))
                if temp:
                    temp1 = temp.group(0).rstrip('}}').lstrip('{{')
                    if (temp1.startswith('head:')):
                        head = temp1[5:]
                        index_of_head.append((row_x,col_x))
                    else:
                        body.append(temp1[5:])
                        indexes_of_body.append((row_x,col_x))
    return function_name, index_of_function,head, index_of_head, body, indexes_of_body,fd

def generate_output(list_objects,index_of_function,  head, index_of_head, body, indexes_of_body, input_file,fname):
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
                value = values[i][h]
                xf_index = sheet.cell_xf_index(indexes_of_body[h][0],indexes_of_body[h][1])
                wtsheet.write(row,col_index,value, style_list[xf_index])
        row += 2
    row -= 1
    for row_index in range(indexes_of_body[0][0] + 1, sheet.nrows, 1):
        if (sheet.rowinfo_map.get(row_index)):
            wtsheet.row(row).height = sheet.rowinfo_map.get(row_index).height
        for col_index in range(sheet.ncols):
            xf_index = sheet.cell_xf_index(row_index, col_index)
            wtsheet.write(row,col_index,sheet.cell(row_index,col_index).value, style_list[xf_index])
        row += 1
    wtbook.save('%s/%s' % (FILE_GENERATE_PATH, fname))
    return message

def generate(filename):
    fname = filename
    try:
        function_name, index_of_function, head, index_of_head, body, indexes_of_body, input_file = fileExtractor(fname)
    except:
        return 'Wrong input file, please check all data'
    else:
        try:
            list_objects = eval('definitions.%s'%function_name)
        except :
            try:
                list_objects = eval(function_name)
            except :
                return 'Definition of function error'
        try:
            len(list_objects)
        except :
            return 'The function you defined returns wrong result (must return a list of objects)'
        message = generate_output(list_objects, index_of_function, head, index_of_head, body, indexes_of_body, input_file,fname)
        return message
    return 'ok'
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list