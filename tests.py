from django.test import TestCase
from report_tool.models import School, Class, Pupil
from datetime import datetime
import settings
import os


class SimpleTest(TestCase):
    def setUp(self):
        self.school = School.objects.create(name = "tung")
        self.class_for_test = Class.objects.create(name = "tung")
        self.pupil = Pupil.objects.create(name = 'tung', school_id = self.school, class_id = self.class_for_test)
    def test_basic_addition(self):
        """
        Tests that 1 + 1 always equals 2.
        """
        self.assertEqual(1 + 1, 2)

    def test_returned_name(self):
        "Upload object should have name same as it's description"
        self.assertEqual(str(self.school), 'tung')
        self.assertEqual(str(self.class_for_test), 'tung')
        self.assertEqual(str(self.pupil), 'tung')