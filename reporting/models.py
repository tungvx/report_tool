# -*- coding: utf-8 -*-
import datetime
from django.db import models
from django import forms
from django.contrib.auth.models import User
from django.contrib.auth import authenticate


class Upload(models.Model):                             #Upload files table in databases
    filename    = models.CharField(max_length=30)
    upload_time = models.DateTimeField('time uploaded')
    description = models.CharField(max_length=255)
    filestore   = models.CharField(max_length=30)
    def __unicode__(self):
        return self.description


class Spreadsheet_report(models.Model): # model to store the information about the spreadsheet used by user
    created_time = models.DateTimeField('time created')
    description = models.CharField(max_length=255)
    spreadsheet_link = models.CharField(max_length=255)
    output_link = models.CharField(max_length=255)
    title = models.CharField(max_length=255)
    def __unicode__(self):
        return self.description

class upload_file_form(forms.Form):                     # Define a simple form for uploading excels file
    description   = forms.CharField(max_length=255,required=True)
    file    = forms.FileField(required=True,)

def handle_uploaded_file(f,location,filename):
                            #Save file upload content to uploaded folder
    fd = open('%s/%s' % (location, str(filename)), 'wb')     #Create new file for write
    for chunk in f.chunks():
        fd.write(chunk)                                 #Write file data
    fd.close()                                          #Close the file

class spreadsheet_report_form(forms.Form):
    description   = forms.CharField(max_length=255,required=True)
    spreadsheet_link = forms.CharField(max_length=255,required=False)

DATABASE_ENGINE_CHOICES = (
    ('odbc', 'odbc'),
    ('access', 'access'),
    ('mssql', 'mssql'),
    ('mysql', 'mysql'),
    ('mxodbc', 'mxodbc'),
    ('mxoracle', 'mxoracle'),
    ('oci8', 'oci8'),
    ('odbc_mssql', 'odbc_mssql'),
    ('vfp', 'vfp'),
    ('sqlite', 'sqlite'),
    ('postgres', 'postgres'),
)

#class User profile
class UserProfile(models.Model):
    user = models.OneToOneField(User)
    database_engine = models.CharField(max_length=200, choices=DATABASE_ENGINE_CHOICES)
    database_name = models.CharField(max_length=200, help_text="Use puns liberally")
    database_user = models.CharField(max_length=200,  blank=True)
    database_password = models.CharField(max_length=200,  blank=True)
    database_host = models.CharField(max_length=200,  blank=True)
    database_port = models.CharField(max_length=200,  blank=True)

class user_profile_form(forms.ModelForm):
    class Meta:
        model = UserProfile
        exclude = ('user',)
