from django.contrib import admin
from django.urls import path, include
from django.conf.urls.static import static
from django.conf import settings
from superadmin.views import*


urlpatterns = [
    
    path('login/', login_page, name='login'),
    path('', login_page),

    path('forget_password/', forget_password, name='forget_password'),
    path('forget_password_send_email/', forget_password_send_email, name='forget_password_send_email'),
    path('forget_password_send_email_check/<uid>/', forget_password_send_email_check,
        name='forget_password_send_email_check'),
    path('update_password/', update_password, name='update_password'),

    path('register/', register_page, name='register_page'),
    path('logout/', logout_page, name='logout'),

    path('adminn/', admin.site.urls),
    path('admin/',  include(('superadmin.urls', 'superadmin'), namespace='superadmin')),
    path('', include(('user.urls', 'user'), namespace='user')),
    
]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
