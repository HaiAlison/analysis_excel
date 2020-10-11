from django.urls import path, include
from . import views
from rest_framework import routers

# router = routers.DefaultRouter()
# router.register('filename', views.ImportFile)
urlpatterns = [
    path('', views.home, name='blog-home'), #views.home redirect to Home function in views.py
    # path('',include(router.urls))
    
]