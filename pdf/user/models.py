from django.db import models
from django.contrib.auth.models import User

# Create your models here.
class MyModel(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    pdf = models.FileField(upload_to='pdf_input')
class Pdf(models.Model):
    pdf = models.FileField(upload_to='pdf_input')

class Storefile(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    pdf = models.FileField(upload_to='pdf_input')
    

