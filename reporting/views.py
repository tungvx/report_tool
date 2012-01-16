from time import time, ctime
import datetime
from django.core.files import File
import os.path
import datetime
from django.core import serializers
from django.http import HttpResponse, HttpResponseRedirect, HttpResponseNotAllowed
from django.core.urlresolvers import reverse
from django.http.multipartparser import FILE
from django.shortcuts import render_to_response
from django.template.loader import render_to_string
from django.template import RequestContext, loader
from django.core.exceptions import *
from django.middleware.csrf import get_token
from django.utils import simplejson
from django.contrib.auth.forms import *
from django.template import Context, loader
from reporting.models import Upload,upload_file_form,handle_uploaded_file
from django.http import HttpResponse,HttpResponseRedirect
import datetime
import reporting.definitions
from django.core.servers.basehttp import FileWrapper
from xlwt.Workbook import Workbook
import xlrd,xlwt
from reporting.report import generate
import mimetypes
import os

SITE_ROOT = os.path.dirname(os.path.realpath(__file__))
UPLOAD = 'upload.html'
FILE_LIST = 'filelist.html'
FILE_UPLOAD_PATH = SITE_ROOT + '/uploaded'
FILE_GENERATE_PATH = SITE_ROOT + '/generated'
FILE_INSTRUCTION_PATH = SITE_ROOT + '/instructions'

def index(request):
    message=None
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

def file_list(request):
    message = None
    file_list = list(Upload.objects.order_by('-upload_time'));
    c = RequestContext(request)
    return render_to_response(FILE_LIST, {'message':message,'file_list':file_list},
                              context_instance = c
                              )
def upload_file(request):
    print FILE_UPLOAD_PATH
    #This function handle upload action
    message=None
    if request.method == 'POST':                # If file fo# rm is  submitted
        form = upload_file_form(request.POST, request.FILES)
        if form.is_valid():                     #Cheking form validate
            f = request.FILES['file']
            fileName, fileExtension = os.path.splitext(f.name);
            if fileExtension!=('.xls'):
                message ='wrong file extension'
            else:
                now = datetime.datetime.now()
                temp = Upload( filestore=str(now.year)+str(now.day)+str(now.month)+str(now.hour)+str(now.minute)+str(now.second)+f.name,filename =f.name,description = request.POST['description'],upload_time=datetime.datetime.now())
                handle_uploaded_file(f, FILE_UPLOAD_PATH,temp.filestore)             #Save file content to uploaded folder
                generator = generate(temp.filestore)
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
                    file_list = list(Upload.objects.order_by('-upload_time'));
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

