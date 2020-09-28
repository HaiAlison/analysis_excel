from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='blog-home'), #views.home redirect to Home function in views.py
]