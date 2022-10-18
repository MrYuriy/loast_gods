from distutils.command.upload import upload
from django.db import models

class File(models.Model):
    filexl = models.FileField(upload_to='filesxl', null=True , blank=True)
