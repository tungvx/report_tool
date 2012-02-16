# -*- coding: utf-8 -*-
import datetime
from django.db import models
from django import forms


class Upload(models.Model):                             #Upload files table in databases
    filename    = models.CharField(max_length=255)
    upload_time = models.DateTimeField('time uploaded')
    description = models.CharField(max_length=255)
    filestore   = models.CharField(max_length=255)
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