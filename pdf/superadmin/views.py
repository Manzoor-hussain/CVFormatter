from base64 import urlsafe_b64decode, urlsafe_b64encode
from email.message import EmailMessage
from django.http import JsonResponse
from django.shortcuts import render, redirect
from django.contrib.auth import login, authenticate, logout
from django.contrib.auth.decorators import login_required
from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.contrib.auth.models import User
from django.contrib.auth.hashers import check_password
from django.middleware.csrf import get_token
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.contrib.auth.tokens import default_token_generator
from django.utils.encoding import force_bytes
from .models import Service, Permission, Myservice, Mypermission
from .serializer import MyypermissionSerializer, ServiceSerializer, PermisstionSerializer, MypermissionSerializer, UserDataSerializer,UserSerializerForCountActivity
from django.core import serializers

import pdb




# Create your views here.
@login_required
def get_index_page(request):
    services_ = Myservice.objects.all()
 

    if request.user.is_superuser:
       
        services_ = Myservice.objects.all()
        return render(request, 'superadmin/index.html',context={'service': services_})
    return redirect('user:dashboard')
 
    return render(request, 'user/index.html',context={'service': services_})
@api_view(['POST','GET'])
def login_page(request):
    if request.method == 'GET':
        user_ = request.user
        if user_.username:
            response = redirect('user:dashboard')
            return response
        else:
            error_message = ''
            return render(request, 'superadmin/auth/login.html', {'error_message': error_message})
        
   
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user_ = authenticate(username=username, password=password)
        
       
      
        if user_:
            login(request, user_)
           

            if user_.is_superuser:
                response = redirect('superadmin:dashboard')
            else:
                response = redirect('user:dashboard')
            response.set_cookie(key='name', value=user_.email)
            response.set_cookie(key='email', value=user_.username)
            
            return response
        else:
            error_message = 'Invalid username or password'
            return render(request, 'superadmin/auth/login.html', {'error_message': error_message})

        #response = render(request, 'superadmin/auth/login.html',context={"errors":"husain"})
        #return response

@api_view(['POST','GET'])
def register_page(request):
    if request.method == 'GET':
        user_ = request.user
        if user_.username:
            response = redirect('user:dashboard')
            return response
        else:
            return render(request, 'superadmin/auth/register.html')
        
    

    if request.method == 'POST':
        name_ = request.POST['full_name']
        email = request.POST['email'].lower().strip()
        password = request.POST['password']
      

        if User.objects.filter(username=email).exists():
            return JsonResponse(status=400, data={'error': 'This email address already registered'})

        uu = User.objects.create(first_name=name_, email=email, username=email)
        uu.set_password(password)
        uu.save()
        try:
            service_instances = Myservice.objects.filter(is_permisstion=True) 
            for service_instance in service_instances:
                Mypermission.objects.create(user=uu, service=service_instance ,is_check=True)
        except:
            pass

      
       
        return JsonResponse(status=200, data={})

    response = render(request, 'superadmin/auth/register.html')
    return response


@login_required
def logout_page(request):
   
    response = redirect('/login')  # Redirect to the desired page after logout
    
    # Set cookies to expire in the past
    response.delete_cookie('name')
    response.delete_cookie('email')
    response.delete_cookie('sessionid')
    response.delete_cookie('csrftoken')
    logout(request)
    # Add more delete_cookie() calls for each cookie you want to remove
    
    return response
    #return redirect('/login')
@api_view(['GET'])
def forget_password(request):
    return render(request, 'superadmin/auth/forgetpassword.html')


class Util:
    @staticmethod
    def send_email(data):
        email = EmailMessage(
            subject=data['subject'],
            body=data['body'],
            from_email="example.com", #os.environ.get('EMAIL_FROM'),
            to=[data['to_email']]
        );
        email.send()


@api_view(['POST'])
def forget_password_send_email(request):
    email = request.POST['email'].lower().strip()

    # if User.objects.filter(username=email).exists():
    #     user=User.objects.filter(username=email).last()
    #     subject = "Password Reset Requested"

    #     return render(request, 'superadmin/auth/reset_password.html',context={'username': email})
    if User.objects.filter(email=email).exists():

        user = User.objects.get(email=email)
        uid = urlsafe_b64encode(force_bytes(user.id))
        # print('Encoded UID', uid)
        # token = PasswordResetTokenGenerator().make_token(user)
        # print('Password Reset Token', token)
        link = 'http://127.0.0.1:8000/forget_password_send_email_check/' + uid + '/'
        print('Password Reset Link', link)
        # Send EMail
        body = 'Click Following Link to Reset Your Password ' + link
        data = {
            'subject': 'Reset Your Password',
            'body': body,
            'to_email': 'manzoorhussain075@gmail.com'
        }
        Util.send_email(data)
        return render(request, 'superadmin/auth/reset_password.html', context={'username': "done"})
    else:
        return render(request, 'superadmin/auth/forgetpassword.html',
                      context={'error': "Your Account Does not Exist pleas"})


@api_view(['GET'])
def forget_password_send_email_check(request, uid):
    id = urlsafe_b64decode(uid)

    user = User.objects.filter(id=id).last()
    username = user.username
    return render(request, 'superadmin/auth/reset_password.html', context={'username': username})


@api_view(['POST'])
def update_password(request):
    new_password = request.POST['password1']
    confirm_password = request.POST['password2']
    email = request.POST['username']

    if new_password != confirm_password:
        return Response(status=400, data={'error': 'Both passwords must match'})

    user = User.objects.filter(username=email).last()
    user.set_password(new_password)
    user.save()
    login(request, user)
    return redirect("login")
@api_view(['POST'])
@login_required
def delete_service(request):
    Myservice.objects.filter(id=request.data['id']).delete()
    return Response(status=200)


@api_view(['POST'])
@login_required
def change_service(request):
    dm_ = Myservice.objects.get(id=request.POST['id'])
    is_permission_ = True
    if dm_.is_permisstion:
        is_permission_ = False
    dm_.is_permisstion = is_permission_
    dm_.save()
    return Response(status=200)


@api_view(['GET'])
@login_required
def get_service_detail(request):
    service_ = Myservice.objects.get(id=request.GET.get('id'))
    data_ = ServiceSerializer(service_)
    data_ = data_.data
    return Response(status=200, data=data_)


@api_view(['POST'])
@login_required
def add_service(request):
    id_ = request.GET.get('id', None)
    title_ = request.POST['title']
    is_permisstion = request.POST.getlist('is_permisstion')
    description = request.POST['description']

   
    is_permisstion = ','.join(is_permisstion)
    if not is_permisstion:
        is_permisstion = 'True'
        
   
  
    if not id_:
        Myservice.objects.create(title=title_, is_permisstion=is_permisstion, description=description)
        return Response(status=200)

    dd = Myservice.objects.get(id=id_)
    dd.description = description
    dd.title = title_
    dd.is_permisstion = is_permisstion
    dd.save()
    return Response(status=200)

@api_view(['GET'])
@login_required
def get_user_detail(request):
    service = Myservice.objects.only('title')
    permission = Mypermission.objects.all()
    users = User.objects.filter(is_superuser=False).exclude(username="hussain").exclude(username="")
    serializers = MypermissionSerializer(permission ,many=True)
    # user=User.objects.get(id=4)
    # permission = Mypermission.objects.all()
    # my=MyypermissionSerializer(permission, many=True)
   

    return render(request, 'superadmin/user_permission.html', context ={"permission":serializers.data,"users":users,"service":service,"final":"sas"})

@api_view(['GET'])
@login_required
def get_user_services(request):
    username_ = request.GET.get('username', None)
    user= User.objects.filter(username=username_).last()
    mypermissions = Mypermission.objects.filter(user=user)
    myservices = Myservice.objects.filter(mypermission__user=user)
   
    json_data = serializers.serialize('json', myservices)   
    return Response(status=200, data=json_data)

@api_view(['POST'])
@login_required
def change_permission(request):
    service_ = Myservice.objects.get(id=request.POST['service'])
    user_ = User.objects.get(username=request.POST['username'])
    ischeck_ = request.POST['ischeck']
    if ischeck_ == 'true':
        permission = Mypermission.objects.create(user=user_, service=service_, is_check=True)
       
    else:
        permission = Mypermission.objects.filter(user=user_, service=service_ , is_check=True)
        permission.delete()

    return Response(status=200)
@api_view(['GET'])
@login_required
def get_user_activity(request):
    username_ = request.GET.get('username', None)
    user_= User.objects.get(username=username_)
    mypermissions = Mypermission.objects.filter(user=user_)
    
    myservices = Myservice.objects.filter(mypermission__user=user_,is_permisstion=True)
    services_ = UserSerializerForCountActivity(myservices, many=True, context={'request':user_})
    services_ = services_.data
   # services_ = UserSerializerForCountActivity(myservices, many=True, context={'request': request.user})
  
        
    return Response(status=200, data=services_)

    
