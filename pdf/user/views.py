from django.http import Http404, HttpResponse
from django.shortcuts import render
from rest_framework.decorators import api_view
from rest_framework.views import APIView
from .serializers import PdfSerializer, StorefileSerializer
from rest_framework.response import Response
from rest_framework import status
from .models import Pdf, Storefile
from django.contrib.auth.decorators import login_required
from django.http import FileResponse
from .services.Expert_resourse import expert_resource_converter
from .services.Joss_Search import joss_search_converter
from .services.Aspion import aspion_converter
from .services.Clarus import clarus_converter
from .services.EdEx import edex_converter
from .services.Sang_Zarrin import sang_zarrin_converter
from .services.FMCG import fmcg_converter
from .services.linum import linum_converter
from .services.Alxander import alexander_steele_converter
from django.http import FileResponse
from django.conf import settings
from superadmin.models import Service, Myservice, Mypermission
import os
import time
import pdb

# Create your views here.
@login_required
def get_index_page(request):
    mypermissions = Mypermission.objects.filter(user=request.user)
    myservices = Myservice.objects.filter(mypermission__user=request.user,is_permisstion=True)

 
    if request.user.is_superuser:
        services_ = Myservice.objects.all()
        return render(request, 'superadmin/index.html',context={'service': services_})
    

    return render(request, 'user/index.html',context={'service': myservices})

@api_view(['GET'])
def index(request):
    file=Pdf.objects.get(id=5)
    return render(request, 'user/index.html', context={"file":file})

@api_view(['GET'])
@login_required
def upload_file(request ,service):
    return render(request, 'user/upload_file.html', context={"service_name":service})
    

@login_required
@api_view(['POST'])
def perform_services(request): 

    data=request.data
    serializer=StorefileSerializer(data=request.data)
    title = request.POST['title']
   
    
 
    if serializer.is_valid():
        serializer.save()
       
        obj=Storefile.objects.filter(user=request.user.id).last()
        path=str(obj.pdf)
        file_path = os.path.join(settings.MEDIA_ROOT, path)
        #/Users/manzoorhussain/Documents/Services/IlovePDF/pdf/media/pdf_output/Expert_Resource.docx
        #ouput_file = "pdf_output/Expert_Resource.docx"
        userfile= str(request.user.id)
        concatenated_str = userfile+"Common_Resource.docx"
        save_path = "pdf_input/"+concatenated_str
        save_path =  os.path.join(settings.MEDIA_ROOT, save_path)
        #file_path_output = os.path.join(settings.MEDIA_ROOT, ouput_file)
      
       
        if title:
            service_name = ("_".join(title.split())+"_Converter").lower()
            output_ = ("pdf_output/"+"_".join(title.split())+"_template.docx").lower()
            file_path_output = os.path.join(settings.MEDIA_ROOT, output_)
            (eval(service_name)(file_path,file_path_output,save_path))
            
        
        return Response(status=200, data=serializer.data)
    return Response(status=400,data=serializer.errors)
 
  
  

api_view(['GET'])
@login_required
def download_docx(request):
   
    
    obj = Storefile.objects.filter(user=request.user.id).last()
    if obj:
        
        id = obj.id
        input_file = str(obj.pdf) 
        input_file_path = os.path.join(settings.MEDIA_ROOT, input_file)
        file_name = "/pdf_input/"
        userfile= str(request.user.id);
        concatenated_str = userfile+"Common_Resource.docx"
        output = "pdf_input/"+concatenated_str
        #output = "pdf_input/Common_Resource.docx"
        file_path = os.path.join(settings.MEDIA_ROOT, output)
        #"/Users/manzoorhussain/Documents/Services/IlovePDF/pdf/media/pdf_input/output_expert_resource.docx"
        filename = os.path.basename(file_path)
        response = HttpResponse(content_type='application/octet-stream')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        # Open the file and write its contents to the response
        with open(file_path, 'rb') as file:
            response.write(file.read())
        
            os.remove(input_file_path)
            os.remove(file_path)
            obj.delete()
        
        
        
            return response
        raise Http404
    response=400
    return response













   
