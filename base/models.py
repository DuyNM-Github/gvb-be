from django.db import models
import uuid

# Create your models here.
class UploadedFile(models.Model):
    session = models.UUIDField(primary_key = False)
    file_id = models.UUIDField(primary_key = False)
    file_name = models.CharField(max_length=255, default='unnamed')