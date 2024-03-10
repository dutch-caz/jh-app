from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name="home"), 
    path('signup', views.signup, name='signup'),
    path('getcoords', views.getcoords, name='getcoords'),
    path('generate_spreadsheet/', views.generate_spreadsheet, name='generate_spreadsheet'),
    path('check_in/', views.check_in, name='check_in'),
    path('success/', views.success, name='success'),
    path('too_far/', views.too_far, name='too_far'),
]
