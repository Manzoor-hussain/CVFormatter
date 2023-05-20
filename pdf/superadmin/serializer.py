from rest_framework import serializers
from .models import Service, Myservice, Mypermission


class ServiceSerializer(serializers.ModelSerializer):
    class Meta:
        model = Myservice
        fields = '__all__'
class PermisstionSerializer(serializers.ModelSerializer):
    class Meta:
        model = Mypermission
        fields = '__all__'