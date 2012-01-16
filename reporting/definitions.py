from report_tool.models import Pupil
from django.contrib.admin.models import LogEntry

def get_ds_hs():
    return Pupil.objects.all()

def get_student_in_class(_class_name):
    return Pupil.objects.filter(class_id__name = _class_name)

def get_admin_log():
    return LogEntry.objects.all()