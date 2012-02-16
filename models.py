# -*- coding: utf-8 -*-
from django.db import models
from django.contrib.auth.models import User
from django.contrib.auth import authenticate
from django import forms

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

DATABASE_ENGINE_CHOICES = (
    ('postgresql_psycopg2', 'postgresql_psycopg2'),
    ('postgresql', 'postgresql'),
    ('mysql', 'mysql'),
    ('sqlite', 'sqlite'),
    ('oracle', 'oracle'),
)

#class User profile
class UserProfile(models.Model):
    user = models.OneToOneField(User)
    database_engine = models.CharField(max_length=200, choices=DATABASE_ENGINE_CHOICES)
    database_name = models.CharField(max_length=200, blank = True)
    database_user = models.CharField(max_length=200,  blank=True)
    database_password = models.CharField(max_length=200,  blank=True)
    database_host = models.CharField(max_length=200,  blank=True)
    database_port = models.CharField(max_length=200,  blank=True)

class user_profile_form(forms.ModelForm):
    sqlite_database_file    = forms.FileField(required=False, help_text="Chose your sqlite database if you use sqlite")
    class Meta:
        model = UserProfile
        exclude = ('user')
        widgets = {
            'database_password': forms.PasswordInput(),
        }