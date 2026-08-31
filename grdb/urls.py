from django.urls import path
from . import views

urlpatterns = [
    path('', views.grdb_dashboard, name='grdb_dashboard'),
]