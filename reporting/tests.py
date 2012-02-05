"""
This file demonstrates writing tests using the unittest module. These will pass
when you run "manage.py test".

Replace this with more appropriate tests for your application.
"""

from django.test import TestCase
from reporting.models import Upload, Spreadsheet_report
from datetime import datetime


class SimpleTest(TestCase):
    def setUp(self):
        self.upload = Upload.objects.create(filename = 'tung.xls', upload_time = datetime.now(), description = "tung", filestore = "tung.xls")
        self.spreadsheet_report = Spreadsheet_report.objects.create(description = 'tung', created_time = datetime.now())
    def test_basic_addition(self):
        """
        Tests that 1 + 1 always equals 2.
        """
        self.assertEqual(1 + 1, 2)

    def test_returned_name(self):
        "Upload object should have name same as it's description"
        self.assertEqual(str(self.upload), 'tung')
        self.assertEqual(str(self.spreadsheet_report), 'tung')