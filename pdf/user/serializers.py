from rest_framework import serializers
from .models import Pdf, Storefile
from superadmin.models import Myservice, Countservices
from datetime import date,  timedelta,datetime



class PdfSerializer(serializers.ModelSerializer):
    class Meta:
        model = Pdf
        fields = '__all__'
class StorefileSerializer(serializers.ModelSerializer):
    class Meta:
        model = Storefile
        fields = '__all__'
class UserSerializerForCount(serializers.ModelSerializer):
    class Meta:
        model = Myservice
        fields = '__all__'

    def to_representation(self, instance):
        ret = super().to_representation(instance)
        user_ = self.context['request'].user
       
        is_active = False
        count = ''
        created_at =''
       

        ubi = Countservices.objects.filter(user=user_, service=instance).last()
        if ubi:
            count = ubi.count
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

           

       
        
      

        # Compare the date with today and yesterday
       
           
        ret['count'] = count
      
        #   ret['time'] = time
     
        return ret
