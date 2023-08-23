from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path("cadastro/", views.cadastro, name="cadastro"),
    path('', views.login, name='login' ),
    path('projeto/', views.projeto, name='projeto'),
    path('logout/', views.custom_logout, name='logout'),
]