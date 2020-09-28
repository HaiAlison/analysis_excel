from .models import Import
from django.forms import ModelForm
class  ImportFile(ModelForm): 
        class Meta:
            model = Import
            fields = ['link_to_specs']
