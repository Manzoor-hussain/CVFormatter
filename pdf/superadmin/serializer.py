from rest_framework import serializers
from .models import Service, Myservice, Mypermission, User




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







#start here
class UserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['id', 'username', 'email']

class MyserviceSerializer(serializers.ModelSerializer):
    class Meta:
        model = Myservice
        fields = ['id', 'title', 'description']

class MyypermissionSerializer(serializers.ModelSerializer):
    user = UserSerializer()  # Nested User serializer
    service = MyserviceSerializer()  # Nested Myservice serializer

    class Meta:
        model = Mypermission
        fields = ['id', 'user', 'service', 'is_check']

class UserDataSerializer(serializers.ModelSerializer):
    mypermissions = MyypermissionSerializer(many=True)  # Nested Mypermission serializer

    class Meta:
        model = User
        fields = ['id', 'username', 'email', 'mypermissions']