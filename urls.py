from django.conf.urls.defaults import patterns, include, url

# Uncomment the next two lines to enable the admin:
from django.contrib import admin
admin.autodiscover()
from django.contrib.auth.views import login, logout

urlpatterns = patterns('',
    url(r'^', include('reporting.urls')),
    # Uncomment the next line to enable the admin:
    url(r'^admin/', include(admin.site.urls)),
    url(r'^accounts/login/$',  login, name='login'),
    url(r'^accounts/profile/$',  'reporting.views.index'),
    url(r'^accounts/logout/$', logout, name='logout'),
    url(r'^accounts/register/$', 'report_tool.views.register', name='register'),
    url(r'^accounts/database_setup/$', 'report_tool.views.setup_database', name='database_setup'),
)
                                                                                                                                                                  