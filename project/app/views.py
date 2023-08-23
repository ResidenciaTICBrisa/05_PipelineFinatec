from django.shortcuts import render
from django.http.response import HttpResponse
from django.contrib.auth.models import User
from django.contrib.auth import authenticate
from django.contrib.auth import login as login_a
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect
from django.shortcuts import redirect
from django.contrib.auth import logout

def cadastro(request):
    if request.method == "GET":
        return render(request, 'cadastro.html')
    else:
        usuario = request.POST.get('usuario')
        senha = request.POST.get('senha')

        user = User.objects.filter(username=usuario).first()

        if user:
            return HttpResponse('Usuário ja existe')

        user = User.objects.create_user(username=usuario, password=senha)
        user.save()

        return HttpResponseRedirect('/')

def login(request):
    if request.method =="GET":
        return render(request, 'login.html')
    else:
        usuario = request.POST.get('usuario')
        senha = request.POST.get('senha')

        user = authenticate(username=usuario, password=senha)

        if user:
            login_a(request, user)
            return HttpResponseRedirect ('projeto/')
        else:
            error_message = 'Usuário ou senha inválido.'
            return render(request, 'login.html', {'error_message': error_message})

@login_required(login_url="/")
def projeto(request):
    # if request.user.is_authenticated:
    #     return HttpResponse('Projetos')
    # else:
    return render(request, 'projeto.html')

def custom_logout(request):
    logout(request)
    return redirect('/')