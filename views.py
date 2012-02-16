from django.shortcuts import render_to_response
from django.template import RequestContext
from reporting.models import handle_uploaded_file
from django.http import HttpResponseRedirect
from django.contrib.auth.decorators import login_required
from django.contrib import auth
from django.contrib.auth.forms import UserCreationForm
from report_tool.models import UserProfile, user_profile_form
from settings import DATABASE_PATH

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
    try:
        current_user = request.user.get_profile()
    except :
        UserProfile.objects.create(user=request.user).save()
        current_user = request.user.get_profile()
    if request.method == 'POST':
        form = user_profile_form(request.POST, instance=current_user)
        if form.is_valid():
            form.save()
        if request.POST['database_engine'] == 'sqlite':
            try:
                f = request.FILES['sqlite_database_file']#get sqlite database file
                handle_uploaded_file(f, DATABASE_PATH,current_user.user.username + '.db') #Save file content to uploaded folder
            except :
                c = RequestContext(request)
                message = 'You must input sqlite file'
                return render_to_response("registration/database_setup.html", {
                    'form' : form, 'message' : message}, context_instance = c)

    else:
        form = user_profile_form(instance=current_user)
    c = RequestContext(request)
    return render_to_response("registration/database_setup.html", {
        'form' : form}, context_instance = c)