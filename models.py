# -*- coding: utf-8 -*-
from django.db import models

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

