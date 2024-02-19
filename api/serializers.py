from rest_framework import serializers
from base.models import UploadedFile


class UploadedFileSerializer(serializers.ModelSerializer):
    class Meta:
        model = UploadedFile
        fields = ('file_id', 'file_name')