from time import time, ctime
from django.core.files import File
import os.path
import datetime
from django.core import serializers
from django.http import HttpResponse, HttpResponseRedirect, HttpResponseNotAllowed
from django.core.urlresolvers import reverse
from django.http.multipartparser import FILE
from django.shortcuts import render_to_response, redirect
from django.template.loader import render_to_string
from django.template import RequestContext, loader
from django.core.exceptions import *
from django.middleware.csrf import get_token
from django.utils import simplejson
from django.contrib.auth.forms import *
from django.template import Context, loader
from reporting.models import Upload,upload_file_form,handle_uploaded_file, Spreadsheet_report, spreadsheet_report_form, UserProfile, user_profile_form
from django.http import HttpResponse,HttpResponseRedirect
import datetime
import reporting.definitions
from django.core.servers.basehttp import FileWrapper
from xlwt.Workbook import Workbook
import xlrd,xlwt
from reporting.report import generate
from reporting.generate_from_spreadsheet import generate_from_spreadsheet
import mimetypes
import os
from urlparse import urlparse, parse_qs
import gdata.service
import settings
from django.contrib.auth.decorators import login_required
from django.contrib import auth
from django.contrib.auth.forms import UserCreationForm
from django import forms

SITE_ROOT = os.path.dirname(os.path.realpath(__file__))
UPLOAD = 'upload.html'
SPREADSHEET_REPORT = 'spreadsheet_report.html'
FILE_LIST = 'filelist.html'
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded'
FILE_GENERATE_PATH = SITE_ROOT + '/generated'
FILE_INSTRUCTION_PATH = SITE_ROOT + '/instructions'

def index(request):
    message= "Welcome to Reporting system"
    t = loader.get_template(os.path.join('index.html'))
    c = RequestContext(request, {
                                 'message':message,
                                }
                       )
    return HttpResponse(t.render(c))

def help(request):
    message=None
    t = loader.get_template(os.path.join('help.html'))
    c = RequestContext(request, {
                                 'message':message,
                                }
                       )
    return HttpResponse(t.render(c))


def download_file(request):
    message = None
    if (request.method == "GET"):
        fname = request.GET['filename']
        path = eval(request.GET['path'])
    try:
        wrapper = FileWrapper( open( '%s/%s' % (path, fname), "r" ) )
        response = HttpResponse(wrapper, mimetype='application/ms-excel')
        response['Content-Disposition'] = u'attachment; filename=%s' % fname
        return response
    except:
        message = 'The file you requested does not exist or is deleted due to time limit!'
        c = RequestContext(request)
        return render_to_response(FILE_LIST, {'message':message},context_instance = c)

@login_required
def file_list(request):
    settings.DATABASES['default']['NAME'] = '/home/tungvx/hellodjango/tung.db'
    print settings.DATABASES['default']['NAME']
    from django.db import connection
    cursor = connection.cursor()
    cursor.execute('select name from sqlite_master')
    a = cursor.fetchall()
    print len(a)
    message = None
    file_list = list(Upload.objects.order_by('-upload_time'))
    spreadsheet_list = list(Spreadsheet_report.objects.order_by('-created_time'))
    c = RequestContext(request)
    return render_to_response(FILE_LIST, {'message':message,'file_list':file_list, 'spreadsheet_list':spreadsheet_list},
                              context_instance = c
                              )

@login_required
def upload_file(request):
    #This function handle upload action
    message=None
    if request.method == 'POST':                # If file fom is  submitted
        form = upload_file_form(request.POST, request.FILES)
        if form.is_valid():                     #Cheking form validate
            f = request.FILES['file']
            fileName, fileExtension = os.path.splitext(f.name);
            if fileExtension!=('.xls'):
                message ='wrong file extension'
            else:
                now = datetime.datetime.now()
                temp = Upload( filestore=str(now.year)+str(now.day)+str(now.month)+str(now.hour)+str(now.minute)+str(now.second)+f.name,filename =f.name,description = request.POST['description'],upload_time=datetime.datetime.now())
                handle_uploaded_file(f, FILE_UPLOAD_PATH,temp.filestore) #Save file content to uploaded folder
                generator = generate(temp.filestore, request.user)
                if generator != "ok":
                    message = generator
                    c = RequestContext(request)
                    os.remove(FILE_UPLOAD_PATH + '/' + temp.filestore)
                    return render_to_response(UPLOAD, {'form':form, 'message':message},
                                              context_instance = c
                                              )
                else:
                    temp.save()       #Save file information into database
                    message="Uploaded successfully. Your uploaded and generated file will be stored shortly. You should download them in the file list page as soon as possible!"
                    c = RequestContext(request)
                    file_list = [temp]
                    return render_to_response(FILE_LIST, {'file_list':file_list, 'message':message},
                              context_instance = c
                              )

        else:   
            message="Error"
            #return HttpResponseRedirect('http://127.0.0.1:8000/admin')
    else:                                   #if file is not submitted that generate the upload form
        form = upload_file_form()
        
    c = RequestContext(request)
    return render_to_response(UPLOAD, {'form':form, 'message':message},
                              context_instance = c
                              )

@login_required
def spreadsheet_report(request): #action to handle create report from google spreadsheet
    message = ''
    if request.method == 'POST': # if the form is submitted
        form = spreadsheet_report_form(request.POST) #get the form
        
        #if the form is valid
        if form.is_valid():
            spreadsheet_key = None

            # get the spreadsheet link from the request
            spreadsheet_link = request.POST.get('spreadsheet_link')

            #get google username
            username = request.POST.get('username')

            #get password of google account
            password = request.POST.get('password')

            # try to extract the key from the spreadsheet link
            try:
                spreadsheet_key = parse_qs(urlparse(spreadsheet_link).query).get('key')[0]
            except :
                message = 'Wrong link'
                c = RequestContext(request)
                return render_to_response(SPREADSHEET_REPORT, {'form':form, 'message':message}, context_instance = c)

            if spreadsheet_key == '' or spreadsheet_key == None: #if the spreadsheet key is empty
                # display error message
                message = 'Please enter the correct spreadsheet link'
                c = RequestContext(request)
                return render_to_response(SPREADSHEET_REPORT, {'form':form, 'message':message}, context_instance = c)
            
            # from the key of the spreadsheet, generate the report
            generator, output_link,title = generate_from_spreadsheet(spreadsheet_key, request.session.get('token'), username, password)

            #if the message is not ok
            if generator != 'ok':
                #render the add report page, and display the error message
                message = generator
                c = RequestContext(request)
                return render_to_response(SPREADSHEET_REPORT, {'form':form, 'message':message}, context_instance = c)
            else:
                #create and save spreadsheet_report object
                now = datetime.datetime.now()
                spreadsheet_report_object = Spreadsheet_report(created_time = now, description = request.POST['description'],spreadsheet_link = spreadsheet_link, output_link = output_link, title = title)
                #uncomment next line to save the report
                spreadsheet_report_object.save()
                message = "Successfully generate the report"
                c = RequestContext(request)
                spreadsheet_list = [spreadsheet_report_object]
                return render_to_response(FILE_LIST, {'message':message,'file_list':file_list, 'spreadsheet_list':spreadsheet_list},
                              context_instance = c
                              )

        else: # if the form is not valid, then raise error
            message = 'Please enter the required fields'
        
    else: #if user want to create new report from spreadsheet
#        request.session['token']=None
#        temp_token = request.GET.get('token')
#        if temp_token:
#            request.session['token'] = temp_token
#        if request.session.get('token') == None:
#            host = request.get_host()
#            next = 'http://' + str(host) + '/add_spreadsheet'
#            scopes = ['https://docs.google.com/feeds/', 'https://spreadsheets.google.com/feeds/']
#            secure = True  # set secure=True to request a secure AuthSub token
#            session = True
#            return redirect(gdata.service.GenerateAuthSubRequestUrl(next, scopes, secure=secure, session=session))
        #if user is authenticated:
        
        form = spreadsheet_report_form()

    c = RequestContext(request)
    return render_to_response(SPREADSHEET_REPORT, {'form':form, 'message':message}, context_instance = c)

#log in action
def login(request):
    user = auth.authenticate(request.POST['username'], request.POST['password'])

#logout
def logout(request):
    auth.logout(request, next_page = '/')

#register function
def register(request):
    form = UserCreationForm()

    if request.method == 'POST':
        form = UserCreationForm(request.POST)
        if form.is_valid():
            new_user = form.save()
            UserProfile.objects.create(user=new_user).save()
            return HttpResponseRedirect("/")

    c = RequestContext(request)
    return render_to_response("registration/register.html", {
        'form' : form}, context_instance = c)

@login_required
def setup_database(request):
    message = None
    current_user = request.user.get_profile()
    if request.method == 'POST':
        form = user_profile_form(request.POST, instance=current_user)
        if form.is_valid():
            form.save()
    else:
        form = user_profile_form(instance=current_user)
    c = RequestContext(request)
    return render_to_response("registration/database_setup.html", {
        'form' : form}, context_instance = c)