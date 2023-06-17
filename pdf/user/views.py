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
from .services.Ashpy import ashbys_converter
from .services.cw_executive import cw_executive_converter
from .services.E_mobility import e_mobility_converter
from .services.fair_recruitment import fair_recruitment_converter
from .services.HBD import hbd_converter
from .services.M2 import m2_partnership_converter
from .services.Timber import timber_seed_format_converter
from .services.Advocate import advocate_group_converter
from .services.Drayton import drayton_converter
from .services.True_Method import true_method_converter
from .services.JLER import jler_converter
from .services.ScaleGenesis import scale_genesis_converter
from .services.FDRecruit import fd_recruit_converter
from .services.Scienta import scienta_converter
from .services.Adway import adway_converter
from .services.FeatherBank import feather_bank_converter
from .services.William import william_blake_converter
from .services.LEO import leo_partner_converter
from .services.Harrington_Morris import harrington_morris_converter
from .services.Mallory import mallory_converter

from django.http import FileResponse
from django.contrib.auth.models import User
from django.conf import settings
from superadmin.models import Service, Myservice, Mypermission, Countservices
from .serializers import UserSerializerForCount
import os
import time
import pdb

# Create your views here.
@login_required
def get_index_page(request):
    mypermissions = Mypermission.objects.filter(user=request.user)
    myservices = Myservice.objects.filter(mypermission__user=request.user,is_permisstion=True)
    services_ = UserSerializerForCount(myservices, many=True, context={'request': request})
    services_ = services_.data
    data=services_
    subset_list = []
    subset_size = 6

    for i in range(0, len(data), subset_size):
        subset = data[i:i+subset_size]
        subset_list.append(subset)
 
    if request.user.is_superuser:
        #username = request.GET.get('username').strip()
        services_ = Myservice.objects.all()
        return render(request, 'superadmin/index.html',context={'service': services_})
    

    return render(request, 'user/index.html',context={'service': services_ ,"data":subset_list})



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
        userfile= str(request.user.id)
        if title:

            title=title.replace("-"," ")
            outputpath = ("_".join(title.split())+".docx").lower()
            concatenated_str = userfile+outputpath
            save_path = "pdf_input/"+concatenated_str
            save_path =  os.path.join(settings.MEDIA_ROOT, save_path)
            service_name = ("_".join(title.split())+"_Converter").lower()
            print("service_name",service_name)
            output_ = ("pdf_output/"+"_".join(title.split())+"_template.docx").lower()
            file_path_output = os.path.join(settings.MEDIA_ROOT, output_)
            (eval(service_name)(file_path,file_path_output,save_path))
            
        
        return Response(status=200, data=serializer.data)
    return Response(status=400,data=serializer.errors)
 
  
  

api_view(['GET'])
@login_required
def download_docx(request):
    title = request.GET.get('title').strip()
   
  
    service_obj = Myservice.objects.get(title=title)
    
    user_ = request.user
    obj = Storefile.objects.filter(user=request.user.id).last()
   
    
    if obj:
        
        id = obj.id
        input_file = str(obj.pdf) 
        input_file_path = os.path.join(settings.MEDIA_ROOT, input_file)
        file_name = "/pdf_input/"
        userfile= str(request.user.id);
        title=title.replace("-"," ")
        outputpath = ("_".join(title.split())+".docx").lower()
        concatenated_str = userfile+outputpath
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
            try:
                counter_obj=Countservices.objects.filter(user=user_, service=service_obj).last()
                count=counter_obj.count
                count= count+1
                counter_obj.count=count
                counter_obj.save()
            except:
                Countservices.objects.create(user=user_, service=service_obj, count=1)
           
        
            os.remove(input_file_path)
            os.remove(file_path)
            obj.delete()
            
        
        
        
            return response
        raise Http404
    response=400
    return response













   
