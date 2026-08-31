from django.urls import path
from . import views

urlpatterns = [
    path('', views.stati_dashboard, name='stati_dashboard'),
]