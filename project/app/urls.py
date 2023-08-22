from django.urls import path
from . import views

urlpatterns = [
    path("cadastro/", views.cadastro, name="cadastro"),
    path('', views.login, name='login' ),
    path('projeto/', views.projeto, name='projeto')
]