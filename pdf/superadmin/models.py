from django.db import models 
from django.contrib.auth.models import User


# Create your models here.
class Service(models.Model):
    user = models.ManyToManyField(User)
    title = models.CharField(max_length=100, null=False, blank=False)
    description = models.TextField(max_length=1000, null=False, blank=False)
    is_permisstion = models.BooleanField(default=False)

class Permission(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    service = models.ForeignKey(Service, on_delete=models.CASCADE)
    is_check = models.BooleanField(default=True)
class Myservice(models.Model):
    user = models.ManyToManyField(User)
    title = models.CharField(max_length=100, null=False, blank=False)
    description = models.TextField(max_length=1000, null=False, blank=False)
    is_permisstion = models.BooleanField(default=False)
class Mypermission(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    service = models.ForeignKey(Myservice, on_delete=models.CASCADE)
    is_check = models.BooleanField(default=True)

  
    
  