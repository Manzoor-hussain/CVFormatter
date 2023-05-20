from rest_framework import serializers
from .models import Pdf, Storefile


class PdfSerializer(serializers.ModelSerializer):
    class Meta:
        model = Pdf
        fields = '__all__'
class StorefileSerializer(serializers.ModelSerializer):
    class Meta:
        model = Storefile
        fields = '__all__'