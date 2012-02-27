try:
    from report_tool.models import Pupil
    from django.contrib.admin.models import LogEntry
except :
    ''

def get_ds_hs():
    return Pupil.objects.all()

def get_student_in_class(_class_name):
    return Pupil.objects.filter(class_id__name = _class_name)

def get_admin_log():
    return LogEntry.objects.all()

try:
    from school.models import *
except :
    print ''

def mark_for_class(request):
#    return Mark.objects.filter(student_id__class_id__name = '6 A1')
    return Mark.objects.filter(subject_id__class_id__id = int(request.session.get('class_id')),term_id__number=int(request.session.get('termNumber')),current=True).order_by('student_id__index','student_id__first_name','student_id__last_name','student_id__birthday')

def student_list(request):
    return Pupil.objects.filter(class_id__id = int(request.session.get('class_id')))

def get_class(request):
    return Class.objects.filter(id = int(request.session.get('class_id')))

def get_class_list(request):
    class_list = Class.objects.filter(year_id__id = int(request.session.get('year_id'))).order_by('name')
    request.session['class_list'] = class_list
    request.session['additional_keys'].append('class_list')
    return class_list

def get_subject_list_by_class(request):
    return Subject.objects.filter(name=request.session.get('subject_name'),class_id__year_id__id = int(request.session.get('year_id'))).order_by('class_id')

def get_subject_list_by_teacher(request):
    return Subject.objects.filter(name=request.session.get('subject_name'),class_id__year_id=int(request.session.get('year_id')),teacher_id__isnull=False).order_by('teacher_id__first_name','teacher_id__last_name')

def get_dh(request):
    termNumber = int(request.session.get('term_number'))
    year_id = int(request.session.get('year_id'))
    type = int(request.session.get('type'))
    school_id = int(request.session.get('school_id'))
    if int(termNumber) < 3:
        if   type == 1:
            danhHieus = TBHocKy.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id, term_id__number=termNumber,
                                               danh_hieu_hk='G').order_by("student_id__index")
        elif type == 2:
            danhHieus = TBHocKy.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id, term_id__number=termNumber,
                                               danh_hieu_hk='TT').order_by("student_id__index")
        elif type == 3:
            danhHieus = TBHocKy.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id, term_id__number=termNumber,
                                               danh_hieu_hk__in=['G', 'TT']).order_by("danh_hieu_hk",
                                                                                      "student_id__index")
    else:
        if   type == 1:
            danhHieus = TBNam.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id, danh_hieu_nam='G').order_by("student_id__index")
        elif type == 2:
            danhHieus = TBNam.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id, danh_hieu_nam='TT').order_by("student_id__index")
        elif type == 3:
            danhHieus=TBNam.objects.filter(student_id__classes__block_id__school_id__id = school_id, student_id__classes__year_id__id=year_id,danh_hieu_nam__in=['G','TT']).order_by("danh_hieu_nam","student_id__index")

    return danhHieus

def get_pupils_no_pass(request):
    type = int(request.session.get('type'))
    school_id = int(request.session.get('school_id'))
    year_id = int(request.session.get('year_id'))
    if   type == 1:
        pupils = TBNam.objects.filter(student_id__classes__block_id__school_id = school_id, student_id__classes__year_id__id=year_id, len_lop=False).order_by("student_id__index")
    elif type == 2:
        pupils = TBNam.objects.filter(student_id__classes__block_id__school_id = school_id, student_id__classes__year_id__id=year_id, thi_lai=True).order_by("student_id__index")
    elif type == 3:
        pupils = TBNam.objects.filter(student_id__classes__block_id__school_id = school_id, student_id__classes__year_id__id=year_id, ren_luyen_lai=True).order_by("student_id__index")

    return pupils