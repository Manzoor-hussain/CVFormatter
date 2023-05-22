from django.contrib import admin

from .models import Service, Permission, Mypermission, Myservice, Countservices

admin.site.register(Service)
admin.site.register(Permission)
admin.site.register(Myservice)
admin.site.register(Mypermission)
admin.site.register(Countservices)

# Register your models here.
