from django.contrib import admin

from .models import Service, Permission, Mypermission, Myservice

admin.site.register(Service)
admin.site.register(Permission)
admin.site.register(Myservice)
admin.site.register(Mypermission)

# Register your models here.
