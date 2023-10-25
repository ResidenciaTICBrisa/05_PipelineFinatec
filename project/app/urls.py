from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

from .views import HomeView

urlpatterns = [
    path('', HomeView.as_view(), name='home'),
    path("cadastro/", views.cadastro, name="cadastro"),
    path('login/', views.login, name='login' ),
    path('projeto/', views.projeto, name='projeto'),
    path('logout/', views.custom_logout, name='logout'),
    # path('login_teste/', views.login_teste, name='login_teste'),
    # path('cadastro_teste/', views.cadastro_teste, name='cadastro_teste'),
    # path('projeto_teste/', views.projeto_teste, name='projeto_teste'),
]