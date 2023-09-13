from django.shortcuts import render
from django.contrib.auth.models import User
from django.contrib.auth import authenticate
from django.contrib.auth import login as login_a
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect
from django.shortcuts import redirect
from django.contrib.auth import logout
from django.contrib.auth.password_validation import validate_password

def cadastro(request):
    if request.method == "GET":
        return render(request, 'cadastro.html')
    else:
        usuario = request.POST.get('usuario')
        senha = request.POST.get('senha')
        senha_confirmacao = request.POST.get('senhaConfirm')
        email = request.POST.get('email')
        first_name = request.POST.get('nome1')
        last_name = request.POST.get('nome2')

        try:
            validate_password(senha, user=User)
        except Exception as e:
            error_messages = e.messages
            return render(request, 'cadastro.html', {'error_messages': error_messages})

        user = User.objects.filter(username=usuario).first()

        if user:
            error_messages = ['Usuário já existe']
            return render(request, 'cadastro.html', {'error_messages': error_messages})
        
        if senha != senha_confirmacao:
            error_messages = ['A senha e a confirmação da senha não coincidem.']
            return render(request, 'cadastro.html', {'error_messages': error_messages})

        user = User.objects.create_user(username=usuario, password=senha, email=email)
        user.is_active = True
        user.first_name = first_name
        user.last_name = last_name
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

# def login_teste(request):
#     if request.method =="GET":
#         return render(request, 'login_teste.html')
#     else:
#         usuario = request.POST.get('usuario')
#         senha = request.POST.get('senha')

#         user = authenticate(username=usuario, password=senha)

#         if user:
#             login_a(request, user)
#             return HttpResponseRedirect ('http://127.0.0.1:8000/projeto/')
#         else:
#             error_message = 'Usuário ou senha inválido.'
#             return render(request, 'login_teste.html', {'error_message': error_message})
        
# def cadastro_teste(request):
#     if request.method == "GET":
#         return render(request, 'cadastro_teste.html')
#     else:
#         usuario = request.POST.get('usuario')
#         senha = request.POST.get('senha')

#         try:
#             validate_password(senha, user=User)
#         except Exception as e:
#             error_messages = e.messages
#             return render(request, 'cadastro_teste.html', {'error_messages': error_messages})

#         user = User.objects.filter(username=usuario).first()

#         if user:
#             error_messages = ['Usuário já existe']
#             return render(request, 'cadastro_teste.html', {'error_messages': error_messages})

#         user = User.objects.create_user(username=usuario, password=senha)
#         user.save()

# @login_required(login_url="/")
# def projeto_teste(request):
#     # if request.user.is_authenticated:
#     #     return HttpResponse('Projetos')
#     # else:
#     return render(request, 'projeto_teste.html')