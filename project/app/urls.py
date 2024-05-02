from django.urls import path
from . import views
from django.contrib.auth import views as auth_views
from .views import user_activity_logs

from .views import HomeView

urlpatterns = [
    path('', HomeView.as_view(), name='home'),
    
    path('login/', views.login, name='login' ),
    path('projeto/', views.projeto, name='projeto'),
    path('logout/', views.custom_logout, name='logout'),
    path('perfil/', views.user_profile, name='user_profile'),
    path('notas/<str:filename>/', views.consultaNotas, name='notas'),
    path('recibos/<str:filename>/', views.download_base64_pdf, name='download_base64_pdf'),
    #path('download-todos-arquivos/<str:filename>/', views.download_todos_arquivos, name='download_todos_arquivos'),
    # path('base/', views.base2, name='base'),-
    
    # path('login_teste/', views.login_teste, name='login_teste'),
    # path('projeto_teste/', views.projeto_teste, name='projeto_teste'),

    path('user-activity-logs/', user_activity_logs, name='user_activity_logs'),
]