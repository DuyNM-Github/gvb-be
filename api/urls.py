from django.urls import path
from api.views import clear_temp, receive_upload_files, process_uploaded_files

urlpatterns = [
    path('files/', receive_upload_files),
    path('process/', process_uploaded_files),
    path('clear/', clear_temp)
]
