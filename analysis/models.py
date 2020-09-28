from django.db import models
from django.utils import timezone
class Import(models.Model):
	link_to_specs = models.FileField(upload_to='media/',max_length=10240)
	date_imported = models.DateTimeField(default=timezone.now) #phần này giống create_at

	# def __str__(self):
	# 	return self.title # cái này là magic method
		
