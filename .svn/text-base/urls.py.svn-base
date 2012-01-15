from django.conf.urls.defaults import patterns, include, url

# Uncomment the next two lines to enable the admin:
from django.contrib import admin
admin.autodiscover()

urlpatterns = patterns('',
    # Examples:
    # url(r'^$', 'report_tool.views.home', name='home'),
    # url(r'^report_tool/', include('report_tool.foo.urls')),
    # Uncomment the admin/doc line below to enable admin documentation:
    # url(r'^admin/doc/', include('django.contrib.admindocs.urls')),

    #(r'^admin/report_tool/upload/$', 'views.index'),
    (r'^upload/add/$', 'views.upload_file'),
    (r'^upload/list/$','views.file_list'),
    (r'^upload/download/$','views.download_file'),
    #(r'^admin/report_tool/uploads/(?P<upload_id>\d+)/$', 'views.detail'),

    # Uncomment the next line to enable the admin:
    url(r'^admin/', include(admin.site.urls)),
    url(r'^$', 'views.file_list'),
    url(r'index$', 'views.file_list'),
    url(r'help$', 'views.help'),

)
