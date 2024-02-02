import os
import datetime
import re
from django.shortcuts import render
from django.contrib.auth.decorators import user_passes_test
from django.contrib.auth.models import User
from django.contrib.auth import authenticate
from django.contrib.auth import login as login_a
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect,HttpResponse
from django.shortcuts import redirect
from django.contrib.auth import logout
from django.contrib.auth.password_validation import validate_password
from django.views.generic import TemplateView
from .models import Template
from .oracle_cruds import consultaPorID
from .new_dev import preenche_planilha,extrair,pegar_caminho
#from .preenche_fub import preencher_fub_teste,consultaID
from .preenche_fundep import preenche_fundep
from .preencheFub import consultaID,preencheFub
from .capa import inserir_round_retangulo
from django.contrib.admin.models import LogEntry
from .models import UserActivity
from django.core.paginator import Paginator

def log_user_activity(user_id, tag, activity):
    UserActivity.objects.create(user_id=user_id, tag=tag, activity=activity)

def convert_datetime_to_string(value):
    if isinstance(value, datetime.datetime):
        return value.strftime('%d/%m/%Y')
    return value

def extract_strings(input_string):
    # Use regular expressions to find the text before and after '@@'
    matches = re.findall(r'(.*?)@@(.*?)@@', input_string)
    
    if matches:
        return tuple(matches[0])
    else:
        
        return (input_string, '')

class HomeView(TemplateView):
    template_name = 'home.html'


# def cadastro(request):
#     if request.method == "GET":
#         return render(request, 'cadastro.html')
#     else:
#         usuario = request.POST.get('usuario')
#         senha = request.POST.get('senha')
#         senha_confirmacao = request.POST.get('senhaConfirm')
#         email = request.POST.get('email')
#         first_name = request.POST.get('nome1')
#         last_name = request.POST.get('nome2')

#         try:
#             validate_password(senha, user=User)
#         except Exception as e:
#             error_messages = e.messages
#             return render(request, 'cadastro.html', {'error_messages': error_messages})

#         user = User.objects.filter(username=usuario).first()

#         if user:
#             error_messages = ['Usuário já existe']
#             return render(request, 'cadastro.html', {'error_messages': error_messages})
        
#         if senha != senha_confirmacao:
#             error_messages = ['A senha e a confirmação da senha não coincidem.']
#             return render(request, 'cadastro.html', {'error_messages': error_messages})

#         user = User.objects.create_user(username=usuario, password=senha, email=email)
#         user.is_active = True
#         user.first_name = first_name
#         user.last_name = last_name
#         user.save()

#         return HttpResponseRedirect('/login/')
    
def login(request):
    if request.method =="GET":
        return render(request, 'login.html')
    else:
        usuario = request.POST.get('usuario')
        senha = request.POST.get('senha')

        user = authenticate(username=usuario, password=senha)

        if user:
            login_a(request, user)
            log_message = f"Acessou o sistema"
            log_user_activity(request.user, "Sistema", log_message)
            
            return HttpResponseRedirect ('/projeto/')
        else:
            error_message = 'Usuário ou senha inválido.'
            
            log_message = f"Tentativa de acesso"
            log_user_activity(request.user, "Sistema", log_message)
            
            return render(request, 'login.html', {'error_message': error_message})

@login_required(login_url="/login/")
def projeto(request):  
    lista_append_db_sql = []
    result = {}
    current_key = None
    mapeamento = None

    codigo = request.POST.get('usuario') # alterar o nome
    template_id = request.POST.get('template')
    # download = request.POST.get('Baixar') # quando puxo na consulta vem vazio
    consultaInicio = request.POST.get('inicio')
    consultaFim = request.POST.get('fim')

    print(f"----------------------\n{codigo}\n{template_id}\n{consultaInicio}\n{consultaFim}\n{request.POST}\n-----------------------")

    print(type(consultaInicio))
    print(consultaFim)
    emptyText = ""

    try:
        db_fin = consultaID(codigo)
    except:
        return render(request,'projeto.html',{
            "templates":Template.objects.all(),
        })  
       
    # nome = Template.objects.get(pk=template_id)
    # nome = Template.objects.get(pk=template_id)
    try:
        nome = Template.objects.get(pk=template_id)
    except:
        return render(request,'projeto.html',{
            "templates":Template.objects.all(),
        })        

    dict_final = {}

    # caminho_pasta_planilhas = pegar_caminho("planilhas")   
    # # caminhoPastaPlanilhasPreenchidas = "../../planilhas_preenchidas/"
    # caminhoPastaPlanilhasPreenchidas = pegar_caminho("planilhas_preenchidas") 
    caminho_pasta_planilhas = "../../planilhas/"    
    caminhoPastaPlanilhasPreenchidas = "../../planilhas_preenchidas/"

    # Obtém o diretório atual do script
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Combina o diretório atual com o caminho para a pasta "planilhas_preenchidas" e o nome do arquivo
       
    if nome.nome_template == "fundep":
        testeCaminhoFundep = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"ModeloFUNDEP.xlsx")
        preenche_planilha(testeCaminhoFundep,dict_final)
    if nome.nome_template == "fub":
        testeCaminhoFub = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"Modelo_Fub.xlsx")
        print(testeCaminhoFub)
        preenche_planilha(testeCaminhoFub,dict_final)
    if nome.nome_template == "opas":
        opas = os.path.join(caminho_pasta_planilhas, "ModeloOPAS.xlsx")
        preenche_planilha(opas,dict_final)
    if nome.nome_template == "fap":
        fap = os.path.join(caminho_pasta_planilhas, "ModeloFAP.xlsx")
        preenche_planilha(fap,dict_final)
    if nome.nome_template == "finep":
        finep = os.path.join(caminho_pasta_planilhas, "ModeloFINEP.xlsx")
        preenche_planilha(finep,dict_final)
        

    file_path = None
    print(f"download{template_id}")
    if template_id == '1':
        keys = ['NomeFavorecido','FavorecidoCPFCNPJ','NomeTipoLancamento',
                'HisLancamento','NumDocPago','DataEmissao','NumChequeDeposito',
                'DataPagamento', 'ValorPago']            
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"planilhaPreenchidaModelo_Fub.xlsx")
            
        # file_path = pegar_caminho('/home/ubuntu/Desktop/05_PipelineFinatec/planilhas_preenchidas/planilhaPreenchidaModelo_Fub.xlsx')
        # data_obj = datetime.strptime(consultaInicio, "%Y-%m-%d")
        # consultaInicio = data_obj.strftime("%d/%m/%Y")
        # data_obj2 = datetime.strptime(consultaFim, "%Y-%m-%d")
        # consultaFim = data_obj2.strftime("%d/%m/%Y")
            
        preencheFub(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),file_path)
        inserir_round_retangulo(file_path,consultaInicio,consultaFim,db_fin)
    elif template_id == '2':
        keys = ['NomeFavorecido','FavorecidoCPFCNPJ','NomeRubrica','NumDocPago',
                'DataEmissao','NumChequeDeposito','DataPagamento', 'ValorPago']            
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"planilhaPreenchidaModeloFUNDEP.xlsx")
            
        #file_path = pegar_caminho('/home/ubuntu/Desktop/05_PipelineFinatec/planilhas_preenchidas/planilhaPreenchidaModeloFUNDEP.xlsx')
        preenche_fundep(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),keys,file_path)
    elif template_id == '3':
        p_opas = os.path.join(caminhoPastaPlanilhasPreenchidas, "ModeloOPAS.xlsx")
        file_path = p_opas
    elif template_id == '4':
        p_fap = os.path.join(caminhoPastaPlanilhasPreenchidas, "ModeloFAP.xlsx")
        file_path = pegar_caminho(p_fap)
    elif template_id == '5':
        p_finep = os.path.join(caminhoPastaPlanilhasPreenchidas, "ModeloFINEP.xlsx")
        file_path = pegar_caminho(p_finep)
    else:
        # Handle cases where 'download' doesn't match any expected values
        return HttpResponse("Invalid download request", status=400)
     
    # Check if the file exists
    if os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/octet-stream')
            #print(f'aaaa{os.path.basename(file_path)}')
            response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
            
            # adicionando log de consulta
            consulta_log = f"Projeto: {codigo} | Modelo: {nome.nome_template} | Inicio da Prest.: {consultaInicio} | Fim da Prest.: {consultaFim}"
            # LogEntry.objects.log_action(user_id=request.user.id, content_type_id=1, object_repr=consulta_log, action_flag=1, change_message="Consulta de prestação de contas")
            
            log_user_activity(request.user, "Consulta",consulta_log)

            return response
    else:
        print("Invalid aaaaaaaaaaa request")

    return render(request,'projeto.html',{
        "templates":Template.objects.all(),
    })

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
#     return render(request, 'projeto_teste.html').


from .models import UserActivity

def is_admin(user):
    return user.is_authenticated and user.is_staff

@user_passes_test(is_admin)
def user_activity_logs(request):
    logs_list = UserActivity.objects.all()
    paginator = Paginator(logs_list, 50)  # Show 50 logs per page

    page = request.GET.get('page')
    logs = paginator.get_page(page)
    
    return render(request, 'user_activity_logs.html', {'logs': logs})