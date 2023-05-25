from rest_framework import serializers
from .models import Countservices, Service, Myservice, Mypermission, User
from datetime import date,  timedelta,datetime





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
class UserSerializerForCountActivity(serializers.ModelSerializer):
    class Meta:
        model = Myservice
        fields = '__all__'

    def to_representation(self, instance):
        ret = super().to_representation(instance)
        
        user_ = self.context['request']

      
        is_active = False
        count = ''
        created_at =''
       

        ubi = Countservices.objects.filter(user=user_, service=instance).last()
      
        if ubi:

          
            count = ubi.count
          
            ret['count'] = count
         
            created_at =ubi.created_at
            date_format = "%B, %d, %Y, %I:%M %p"           
            formatted_date = created_at.strftime('%Y-%m-%d')
            hour = created_at.hour
            minute = created_at.minute
            second = created_at.second
            time=f"Time: {hour}:{minute:02d}:{second:02d}"
            current_time = datetime.now()
            formatted_today = current_time.strftime("%Y-%m-%d")
            date_object = datetime.strptime(formatted_date, '%Y-%m-%d').date()
            today_object = datetime.strptime(formatted_today, '%Y-%m-%d').date()

            # Calculate the time difference between the two dates
            time_difference = today_object - date_object
            
            if formatted_date == formatted_today:
                ret['created_at'] = "Today"
            elif time_difference.days == 1:
                ret['created_at'] = "Yesterday"
            else:
                ret['created_at'] = formatted_date
            ret['time'] = time
        return ret

           









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