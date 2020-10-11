from rest_framework import serializers
from .models import Import

class ImportSerializer(serializers.HyperlinkedModelSerializer):

    class Meta:
        model = Import
        fields = '__all__'