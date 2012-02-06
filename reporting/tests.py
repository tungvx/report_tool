from django.test import TestCase
from reporting.models import Upload, Spreadsheet_report
from datetime import datetime
from extract_information import get_list_of_object, extract_information
from generate_from_spreadsheet import upload_result


class SimpleTest(TestCase):
    def setUp(self):
        self.upload = Upload.objects.create(filename = 'tung.xls', upload_time = datetime.now(), description = "tung", filestore = "tung.xls")
        self.spreadsheet_report = Spreadsheet_report.objects.create(description = 'tung', created_time = datetime.now())

    def test_returned_name(self):
        "Upload object should have name same as it's description"
        self.assertEqual(str(self.upload), 'tung')
        self.assertEqual(str(self.spreadsheet_report), 'tung')

    def test_get_list_of_object(self):
        #test if the function get_list_of_object is correct
        message, objects_list = get_list_of_object('Upload.objects.all()', [(1,2)])
        self.assertEqual(message, 'ok') #check if the returned message is 'ok;

        #check if the returned objects list is the correct list
        self.assertEqual(objects_list[0],self.upload)
        self.assertEqual(len(objects_list),1)

        #test for exception, when both argument of this function is empty:
        message, objects_list = get_list_of_object('',[])
        self.assertEqual(message, 'The data function should be specify!')
        self.assertEqual(objects_list, [])

        #test if the function is not correct, then the correct message should be returned
        message, objects_list = get_list_of_object('toilatung', [(1,2)])
        self.assertEqual(message, 'Definition of data function error at cell C2')
        self.assertEqual(objects_list, [])

        #test if the function if correct, but returned value of object_list is not appropriate
        message, objects_list = get_list_of_object('Upload.objects', [(1,2)])
        self.assertEqual(message, 'The function you defined returns wrong result (must return a list of objects):cell C2')
        self.assertEqual(objects_list, [])

    def test_extract_information_function(self):
        index_of_function = []
        index_of_head = []
        body = []
        indexes_of_body = []
        index_of_excel_function = []
        excel_function = []
        other_info = []
        index_of_other_info = []

        #call function to test function_name extraction
        function_name, head = extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, '#<function()>', 1, 2, other_info, index_of_other_info)
        self.assertEqual(function_name, 'function()') #test if the function_name is extracted correctly
        self.assertEqual(head, '') #test if the head is extracted correctly
        self.assertEqual(index_of_function, [(1,2)]) #test if the the index of function_name is assigned correctly
        self.assertEqual(index_of_head,[]) #the index of head should be empty
        self.assertEqual(body, []) #the body should be empty
        self.assertEqual(indexes_of_body, []) #the index of body should be empty
        self.assertEqual(index_of_excel_function, []) #the index of excel function should be empty
        self.assertEqual(excel_function, []) #the excel function list should be empty
        self.assertEqual(other_info, []) #other information should be empty
        self.assertEqual(index_of_other_info, []) #index of other information should be empty

        #test for head extraction
        function_name, head = extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, '{{head:head}}', 1, 2, other_info, index_of_other_info)
        self.assertEqual(function_name, '') #test if the function_name is extracted correctly
        self.assertEqual(head, 'head') #test if the head is extracted correctly
        self.assertEqual(index_of_function, [(1,2)]) #test if the the index of function_name is assigned correctly
        self.assertEqual(index_of_head,[(1,2)]) #the index of head should be [(1,2)]
        self.assertEqual(body, []) #the body should be empty
        self.assertEqual(indexes_of_body, []) #the index of body should be empty
        self.assertEqual(index_of_excel_function, []) #the index of excel function should be empty
        self.assertEqual(excel_function, []) #the excel function list should be empty
        self.assertEqual(other_info, []) #other information should be empty
        self.assertEqual(index_of_other_info, []) #index of other information should be empty

        #test for body extraction
        function_name, head = extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, '{{body:body}}', 1, 2, other_info, index_of_other_info)
        self.assertEqual(function_name, '') #test if the function_name is extracted correctly
        self.assertEqual(head, '') #test if the head is extracted correctly
        self.assertEqual(index_of_function, [(1,2)]) #test if the the index of function_name is assigned correctly
        self.assertEqual(index_of_head,[(1,2)]) #the index of head should be [(1,2)]
        self.assertEqual(body, ['body']) #the body should be ['body']
        self.assertEqual(indexes_of_body, [(1,2)]) #the index of body should be empty
        self.assertEqual(index_of_excel_function, []) #the index of excel function should be empty
        self.assertEqual(excel_function, []) #the excel function list should be empty
        self.assertEqual(other_info, []) #other information should be empty
        self.assertEqual(index_of_other_info, []) #index of other information should be empty

        #test for excel_function extraction
        function_name, head = extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, ':= "{{body:body2}}" + "tung"', 1, 2, other_info, index_of_other_info)
        self.assertEqual(function_name, '') #test if the function_name is extracted correctly
        self.assertEqual(head, '') #test if the head is extracted correctly
        self.assertEqual(index_of_function, [(1,2)]) #test if the the index of function_name is assigned correctly
        self.assertEqual(index_of_head,[(1,2)]) #the index of head should be [(1,2)]
        self.assertEqual(body, ['body', 'body2']) #the body should be ['body', 'body2']
        self.assertEqual(indexes_of_body, [(1,2), (1,2)]) #the index of body should be empty
        self.assertEqual(index_of_excel_function, [(1,2)]) #the index of excel function should be [(1,2)]
        self.assertEqual(excel_function, [':= "{{body:body2}}" + "tung"']) #the excel function list should be correct
        self.assertEqual(other_info, []) #other information should be empty
        self.assertEqual(index_of_other_info, []) #index of other information should be empty

        #test for other information extraction:
        function_name, head = extract_information(index_of_function, index_of_head, body, indexes_of_body,
                        index_of_excel_function, excel_function, 'tung', 1, 2, other_info, index_of_other_info)
        self.assertEqual(function_name, '') #test if the function_name is extracted correctly
        self.assertEqual(head, '') #test if the head is extracted correctly
        self.assertEqual(index_of_function, [(1,2)]) #test if the the index of function_name is assigned correctly
        self.assertEqual(index_of_head,[(1,2)]) #the index of head should be [(1,2)]
        self.assertEqual(body, ['body', 'body2']) #the body should be ['body', 'body2']
        self.assertEqual(indexes_of_body, [(1,2), (1,2)]) #the index of body should be empty
        self.assertEqual(index_of_excel_function, [(1,2)]) #the index of excel function should be [(1,2)]
        self.assertEqual(excel_function, [':= "{{body:body2}}" + "tung"']) #the excel function list should be correct
        self.assertEqual(other_info, ['tung']) #other information should be correct
        self.assertEqual(index_of_other_info, [(1,2)]) #index of other information should be correct

    #function to test upload_result function
    def test_upload_result(self):
        #test for wrong email and password
        message,output_link = upload_result('20121210290.xls','', 'username', 'password')
        self.assertEqual(message, 'Wrong email or password!') #the message returned should be correct
        self.assertEqual(output_link, '') #the returned output_link should be empty

        #test for wrong filename:
        message, output_link = upload_result('noname.xls','','toilatungfake1', 'toilatung')
        self.assertEqual(message, 'Invalid file!')
        self.assertEqual(output_link, '')

        #test the success of function if the parameters are correct
        message, output_link = upload_result('20121210290.xls','','toilatungfake1', 'toilatung')
        self.assertEqual(message, 'ok')
        

