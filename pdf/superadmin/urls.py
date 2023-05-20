from django.contrib import admin
from django.urls import path, include
from .views import *


urlpatterns = [
    path('dashboard/', get_index_page, name='dashboard'),
    path('delete-service/', delete_service, name='delete_service'),
    path('change_service/', change_service, name='change_bank'),
    path('get_service_detail/', get_service_detail, name='get_service_detail'),
    path('add-service/', add_service, name='add_service'),

    path('get_user_detail/', get_user_detail, name='get_user_detail'),

    path('get-user-services/', get_user_services, name='get_user_services'),
    path('change-permission/', change_permission, name='change_permission'),
    
    


    
]
