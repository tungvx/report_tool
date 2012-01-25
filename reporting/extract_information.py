import re

#this function is used for extracting information from a string input value
def extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, value, row_x, col_x):
    function_name = ''
    head = ''
    temp = re.search('#<.*?>', unicode(value)) #if the cell contains the function which returns the data
    if temp:
        function_name = (temp.group(0).rstrip('>').lstrip('#<')) #remove > at the right and #< at the left
        index_of_function.append((row_x, col_x)) #stores the index of this function
    else:
        temp = re.findall('{{.*?}}', unicode(value)) # find all the specified fields of data
        if temp: #if yes
            for temp1 in temp: #iterating all of the fields
                temp1 = temp1.rstrip('}}').lstrip('{{') # remove tags to get attributes
                if (temp1.startswith('head:')): #if the field is the header
                    head = temp1[5:] #remove head:
                    index_of_head.append((row_x, col_x)) #stores the location of the header
                else:
                    body.append(temp1[5:]) #else the field is the body
                    indexes_of_body.append((row_x, col_x)) #stores the location of the body
        if value.startswith(":="):
            excel_function.append(value) #strores the value of the cell contain the specified excel function
            index_of_excel_function.append((row_x, col_x)) #store index of above excel function

    return function_name, head