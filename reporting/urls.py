from django.conf.urls.defaults import patterns, include, url

# Uncomment the next two lines to enable the admin:
from django.contrib import admin
admin.autodiscover()

urlpatterns = patterns('reporting.views',
    # Examples:
    # url(r'^$', 'report_tool.views.home', name='home'),
    # url(r'^report_tool/', include('report_tool.foo.urls')),
    # Uncomment the admin/doc line below to enable admin documentation:
    # url(r'^admin/doc/', include('django.contrib.admindocs.urls')),

    #(r'^admin/report_tool/upload/$', 'views.index'),
    url(r'^add/$', 'upload_file',name='upload_file'),
    url(r'^add_spreadsheet/$', 'spreadsheet_report',name='spreadsheet_report'),
    url(r'^list/$','file_list',name='file_list'),
    url(r'^download/$','download_file',name='download_file'),
    #(r'^admin/report_tool/uploads/(?P<upload_id>\d+)/$', 'views.detail'),

    # Uncomment the next line to enable the admin:
    url(r'^admin/', include(admin.site.urls)),
    url(r'^$', 'index'),
    url(r'index$', 'index'),
    url(r'help$', 'help'),
)
