from django.urls import path
from . import views

urlpatterns = [
    path('', views.rfq_dashboard, name='rfq_dashboard'),
    path("cleaner/", views.rfq_cleaner_view, name="rfq_cleaner"),
    path("generate-word/", views.generate_word, name="generate_word"),
    path("customer-details/",views.customer_details,name="customer_details"),

]