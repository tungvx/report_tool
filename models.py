# -*- coding: utf-8 -*-
import datetime
from django.db import models
from django import forms
from settings import FILE_UPLOAD_PATH
from django.contrib.auth.models import User
from django.contrib.auth import authenticate


class Upload(models.Model):                             #Upload files table in databases
    filename    = models.CharField(max_length=30)
    upload_time = models.DateTimeField('time uploaded')
    description = models.CharField(max_length=255)
    filestore   = models.CharField(max_length=30)
    def __unicode__(self):
        return self.description


class upload_file_form(forms.Form):                     # Define a simple form for uploading excels file
    description   = forms.CharField(max_length=30,required=True)
    file    = forms.FileField(required=True,)

def handle_uploaded_file(f,location,filename):
                            #Save file upload content to uploaded folder
    fd = open('%s/%s' % (location, str(filename)), 'wb')     #Create new file for write
    for chunk in f.chunks():
        fd.write(chunk)                                 #Write file data
    fd.close()                                          #Close the file
class School(models.Model):
    name    = models.CharField(max_length= 25)
    def __unicode__(self):
        return str(self.name)
class Class(models.Model):
    name    = models.CharField(max_length= 25)
    def __unicode__(self):
        return str(self.name)
class Pupil(models.Model):
    name    = models.CharField(max_length=25)
    school_id  = models.ForeignKey(School)
    class_id   =    models.ForeignKey(Class)
    def __unicode__(self):
        return str(self.name)

