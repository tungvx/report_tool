from report_tool.models import Upload,Class,Pupil,School
from django.contrib import admin

class UploadAdmin(admin.ModelAdmin):
    fieldsets = [
        (None,               {'fields': ['filename']}),
        (None,               {'fields': ['description']}),
        ('Date information', {'fields': ['upload_time'], 'classes': ['collapse']}),
    ]
    list_display = ('filename', 'upload_time', 'description')

admin.site.register(Upload, UploadAdmin)
admin.site.register(Class)
admin.site.register(School)
admin.site.register(Pupil)