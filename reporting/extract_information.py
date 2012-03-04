import re
import xlwt
from reporting.models import Upload, Spreadsheet_report
import django
import definitions

try:
    import sqlite3
    import psycopg2
    import psycopg2.extras
    import MySQLdb
except :
    ''

try:
    from report_tool.settings import DATABASE_PATH
except :
    ''

#this function is used for extracting information from a string input value
def extract_information(index_of_function, index_of_group, body, indexes_of_body,
                        index_of_excel_function, excel_function, value, row_x, col_x, other_info, index_of_other_info,
                        body_input, indexes_of_body_input, head, index_of_head, head_input, index_of_head_input,
                        foot, index_of_foot, foot_input, index_of_foot_input,
                        once, index_of_once, once_input, index_of_once_input, group, reserve_postions):
    function_name = ''
    value = unicode(value)
    temp = re.search('#<.*?>', value) #if the cell contains the function which returns the data
    if temp:
        function_name = (temp.group(0).rstrip('>').lstrip('#<')) #remove > at the right and #< at the left
        index_of_function.append((row_x, col_x)) #stores the index of this function
        if (row_x, col_x) not in reserve_postions:
            reserve_postions.append((row_x, col_x))
    else:
        temp = re.findall('{{.*?}}', unicode(value)) # find all the specified fields of data
        if temp: #if yes
            for temp1 in temp: #iterating all of the fields
                temp1 = temp1.rstrip('}}').lstrip('{{') # remove tags to get attributes
                if (temp1.startswith('group')): #if the field is the group
                    temp_group = temp1[5:] #remove group:
                    group_key = temp_group[:temp_group.index(':')]
                    group[group_key] = temp_group[temp_group.index(':') + 1:]
                    index_of_group[group_key] = (row_x, col_x) #stores the location of the group
                elif (temp1.startswith('head')): #if the field is the group:
                    temp_head = temp1[4:] #else the field is the head
                    head_key = temp_head[:temp_head.index(':')]
                    if not head.get(head_key):
                        head[head_key] = []
                        index_of_head[head_key] = []
                        index_of_head_input[head_key] = []
                        head_input[head_key] = []
                    head[head_key].append(temp_head[temp_head.index(':') + 1:])
                    index_of_head[head_key].append((row_x, col_x)) #stores the location of the head
                    if (row_x, col_x) not in index_of_head_input.get(head_key):
                        head_input[head_key].append(value)
                        index_of_head_input[head_key].append((row_x, col_x))
                elif (temp1.startswith('foot')): #if the field is the footer
                    temp_foot = temp1[4:]
                    foot_key = temp_foot[:temp_foot.index(':')]
                    if not foot.get(foot_key):
                        foot[foot_key] = []
                        index_of_foot[foot_key] = []
                        index_of_foot_input[foot_key] = []
                        foot_input[foot_key] = []
                    foot[foot_key].append(temp_foot[temp_foot.index(':') + 1:])
                    index_of_foot[foot_key].append((row_x, col_x))
                    if (row_x, col_x) not in index_of_foot_input.get(foot_key):
                        foot_input[foot_key].append(value)
                        index_of_foot_input[foot_key].append((row_x, col_x))
                elif (temp1.startswith('once:')): #if the field is the footer
                    if (row_x, col_x) not in index_of_once_input:
                        once_input.append(value) # add value to foot array
                        index_of_once_input.append((row_x, col_x)) #also store index of foot

                    once.append(temp1[5:]) #store the field of foot
                    index_of_once.append((row_x, col_x))
                else:
                    if (row_x, col_x) not in indexes_of_body_input:
                        body_input.append(value)
                        indexes_of_body_input.append((row_x, col_x))
                    body.append(temp1) #else the field is the body
                    indexes_of_body.append((row_x, col_x)) #stores the location of the body
            if value.startswith(":="):
                excel_function.append(value) #strores the value of the cell contain the specified excel function
                index_of_excel_function.append((row_x, col_x)) #store index of above excel function
            if (row_x, col_x) not in reserve_postions:
                reserve_postions.append((row_x, col_x))
        else:
            other_info.append(value) #store other information
            index_of_other_info.append((row_x,col_x))#store the index of other information

    return function_name

#function to get a list of objects containing the data
def get_list_of_object(function_name, index_of_function, request):
    if function_name == '':
        return 'ok', []
    #try to get list of objects from definitions.py file, or execute the fuction directly
    try:
        list_objects = eval('definitions.%s' %function_name)
    except :
        try:
            list_objects = eval(function_name)
        except :
            print 'error'
    #if the list is not empty, then return the list
    try:
        if len(list_objects) >= 0:
            return 'ok', list_objects
    except :
        print 'error'

    try:
        current_user = request.user.get_profile() #get user profile
    except :
        return 'You must set up you database!', []

    try:
        database_engine = current_user.database_engine #get database engine
    except :
        try:
            return 'Data specification error at cell ' + xlwt.Utils.rowcol_to_cell(index_of_function[0][0],index_of_function[0][1]), []
        except :
            return 'The data function must be specified!', []

    if database_engine == 'sqlite':
        #connect to sqlite database
        try:
            connection = sqlite3.connect(database = DATABASE_PATH + '/' + user.username + '.db')
            connection.row_factory = dict_factory
            cursor = connection.cursor()
        except :
            return "Wrong database file!", []
    elif database_engine == 'mysql':
        #connect to mysql
        try:
            connection = MySQLdb.connect (host = current_user.database_host,
                             user = current_user.database_user,
                             passwd = current_user.database_password,
                             db = current_user.database_name)
            cursor = connection.cursor (MySQLdb.cursors.DictCursor)
        except:
            return 'Wrong database settings!', []
    elif database_engine == 'postgresql':
        try:
            print "dbname='%s' user='%s' host='%s' password='%s'" %(current_user.database_name, current_user.database_user, current_user.database_host, current_user.database_password)
            connection = psycopg2.connect("dbname='%s' user='%s' host='%s' password='%s'" %(current_user.database_name, current_user.database_user, current_user.database_host, current_user.database_password));
            cursor = connection.cursor(cursor_factory=psycopg2.extras.DictCursor)
        except :
            return 'Wrong database settings',[]
        
    try:
        cursor.execute(function_name)
        list_objects = cursor.fetchall()
    except :
            try:
                return 'Query syntax error at cell ' + xlwt.Utils.rowcol_to_cell(index_of_function[0][0],index_of_function[0][1]), []
            except :
                return 'The query must be specified!', []

    #close connection and rollback:
    connection.close()
    
    return 'ok', list_objects


def dict_factory(cursor, row):
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d