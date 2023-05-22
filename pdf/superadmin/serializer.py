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
class MypermissionSerializer(serializers.ModelSerializer):
    class Meta:
        model = Mypermission
        fields = '__all__'

class MyserviceSerializer(serializers.ModelSerializer):
    permissions = MypermissionSerializer(many=True)

    class Meta:
        model = Myservice
        fields = ['id', 'user', 'title', 'description', 'is_permisstion', 'permissions']

    def create(self, validated_data):
        permissions_data = validated_data.pop('permissions')
        myservice = Myservice.objects.create(**validated_data)
        for permission_data in permissions_data:
            Mypermission.objects.create(myservice=myservice, **permission_data)
        return myservice