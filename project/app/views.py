from django.shortcuts import render
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
from .preenche_fub import preencher_fub_teste
import os
import datetime
import re

from .capa import inserir_round_retangulo
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

        return HttpResponseRedirect('/login/')
    
def login(request):
    if request.method =="GET":
        return render(request, 'login.html')
    else:
        usuario = request.POST.get('usuario')
        senha = request.POST.get('senha')

        user = authenticate(username=usuario, password=senha)

        if user:
            login_a(request, user)
            return HttpResponseRedirect ('/projeto/')
        else:
            error_message = 'Usuário ou senha inválido.'
            return render(request, 'login.html', {'error_message': error_message})

@login_required(login_url="/login/")
def projeto(request):
    # if request.user.is_authenticated:
    #     return HttpResponse('Projetos')
    # else:
        
        global tabe
        lista_append_db_sql = []
        result = {}
        current_key = None
        mapeamento = None
        coduaigo = request.POST.get('usuario')
        template_id = request.POST.get('template')
        download = request.POST.get('Baixar')
        data1 = request.POST.get('inicio')
        data2 = request.POST.get('fim')
        print(type(data1))
        print(data2)
        try:
            db_fin = consultaPorID(coduaigo)
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
        mapeamento = nome.mapeamento
        #print(nome.mapeamento)
        attribute_names = [
                #"id_mapeamento",
                "codigo",
                "nome",
                "saldo",
                "data_assinatura",
                "data_vigencia",
                "data_encerramento",
                "tipo_contrato",
                "instituicao_executora",
                "processo",
                "subprocesso",
                "cod_proposta",
                "proposta",
                "objetivos",
                "valor_aprovado",
                "nome_tp_controle_saldo",
                "grupo_gestores",
                "gestor_resp",
                "coordenador",
                "procedimento_compra",
                "tab_frete",
                "tab_diarias",
                "custo_op",
                "nome_financiador",
                "departamento",
                "situacao",
                "banco",
                "agencia_bancaria",
                "conta_bancaria",
                "centro_custo",
                "conta_caixa",
                "categoria_projeto",
                "cod_convenio_conta",
                "cod_status",
                "ind_sub_projeto",
                "tipo_custo_op",
                "projeto_mae",
                "id_coordenador",
                "id_financiador",
                "id_instituicao",
                "id_departamento",
                "nome_instituicao",
                "id_instituicao_executora",
                "id_tipo"
            ]
       
            # Check for non-empty attributes and print their names
        for attribute_name in attribute_names:
                attribute_value = getattr(mapeamento, attribute_name)
                if attribute_value:
                    lista_append_db_sql.append(f"{attribute_value}@@{attribute_name}@@")
                    
        print(lista_append_db_sql)
        print('\n')
        print(mapeamento.id_mapeamento)
        print('\n')
        print(mapeamento.data_vigencia)
        output = []
        result = {}
        current_key = None
        current_subkey = None

        for line in lista_append_db_sql:
            parts = line.strip().split(";")
            i = 0
        
            while i < len(parts):
                if i + 2 < len(parts):
                    key = parts[i]
                    subkey = parts[i + 1]
                    subsubkey = parts[i + 2]
                    value = extrair(parts)
                    #print(value)
                    
                    if key == current_key:
                        
                        result[key].append((subkey,f"{subsubkey}@@{value[0].upper()}@@"))
                        
                    
                    else:
                        # If the key is different, create a new list
                        current_key = key
                        if key in result:
                            result[key].append((subkey,f"{subsubkey}@@{value[0].upper()}@@"))
                        else:
                            result[key]= [(subkey, f"{subsubkey}@@{value[0].upper()}@@")]
                i += 3
                
            output_dict = {key: value for key, value in result.items()}


        #print(output_dict)

        # for key, value_list in output_dict.items():
        #     for i, (position, template) in enumerate(value_list):
        #         placeholder = None
        #         if "'" in template:
        #             placeholder = template.split("'")[1]
        #         if placeholder in db_fin:
        #             value_to_insert = db_fin[placeholder]
        #             # Convert datetime objects to strings if necessary
        #             if isinstance(value_to_insert, datetime.datetime):
        #                 value_to_insert = value_to_insert.strftime('%Y-%m-%d')  
        #             if value_to_insert is not None:
        #             # Replace the template with the actual value
        #                 value_list[i] = (position, template.replace(f"'{placeholder}'", value_to_insert))

        # Crie um novo dicionário para armazenar os resultados
        novo_dicionario = {}

    # Itere sobre o primeiro dicionário
        for chave, lista_de_tuplas in output_dict.items():
            nova_lista_de_tuplas = []
            for tupla in lista_de_tuplas:
                chave_do_segundo_dicionario = tupla[1]
                #print(chave_do_segundo_dicionario)
                #print(type(chave_do_segundo_dicionario))
                string_before, string_between = extract_strings(chave_do_segundo_dicionario)
                valor_do_segundo_dicionário = db_fin.get(string_between, '')
                #print(valor_do_segundo_dicionário)
                valor_formatado = convert_datetime_to_string(valor_do_segundo_dicionário)
                #print(valor_formatado)
                #nova_tupla = (tupla[0],f"{strings[0]} {valor_formatado}")
                nova_tupla = (tupla[0],f"{string_before}{valor_formatado}")
                nova_lista_de_tuplas.append(nova_tupla)
            novo_dicionario[chave] = nova_lista_de_tuplas

        #print(novo_dicionario)


        dict_final = {}
        for key, values in novo_dicionario.items():
            combined_values = {}
            for item in values:
                if item[0] in combined_values:
                    combined_values[item[0]] += ' ' + item[1]  # Add a space before appending
                else:
                    combined_values[item[0]] = item[1]
    
            dict_final[key] = [(k, v) for k, v in combined_values.items()]

        #print(dict_final)
        tabe = None
        if nome.nome_template == "fundep":
            tabe = preenche_planilha("planilhas/ModeloFUNDEP.xlsx",dict_final)
        if nome.nome_template == "fub":
            tabe = preenche_planilha("planilhas/Modelo_Fub.xlsx",dict_final)
        if nome.nome_template == "opas":
            tabe = preenche_planilha("planilhas/ModeloOPAS.xlsx",dict_final)
        if nome.nome_template == "fap":
            tabe = preenche_planilha("planilhas/ModeloFAP.xlsx",dict_final)
        if nome.nome_template == "finep":
            tabe = preenche_planilha("planilhas/ModeloFINEP.xlsx",dict_final)
        
       
        


    
        file_path = None
        print(f"download{template_id}")
        if template_id == '1':
            keys = ['NOME_FAVORECIDO','CNPJ_FAVORECIDO','TIPO_LANCAMENTO','HIS_LANCAMENTO','DATA_EMISSAO','DATA_PAGAMENTO', 'VALOR_PAGO']
            file_path = pegar_caminho('planilhas_preenchidas/planilhas/Modelo_Fub.xlsx')
            # data_obj = datetime.strptime(data1, "%Y-%m-%d")
            # data1 = data_obj.strftime("%d/%m/%Y")
            # data_obj2 = datetime.strptime(data2, "%Y-%m-%d")
            # data2 = data_obj2.strftime("%d/%m/%Y")
            preencher_fub_teste(coduaigo,convert_datetime_to_string(data1),convert_datetime_to_string(data2),keys,file_path)
            inserir_round_retangulo(file_path,data1,data2,db_fin)
        elif template_id == '2':
           
            file_path = pegar_caminho('planilhas_preenchidas/planilhas/ModeloFUNDEP.xlsx')
        elif template_id == '3':
            
            file_path = pegar_caminho('planilhas_preenchidas/planilhas/ModeloOPAS.xlsx')
        elif template_id == '4':
            
            file_path = pegar_caminho('planilhas_preenchidas/planilhas/ModeloFAP.xlsx')
        elif template_id == '5':
            
            file_path = pegar_caminho('planilhas_preenchidas/planilhas/ModeloFINEP.xlsx')
        else:
            # Handle cases where 'download' doesn't match any expected values
            return HttpResponse("Invalid download request", status=400)
        #print(file_path)
        # Check if the file exists
     
        #print(os.path.exists(file_path))
        if os.path.exists(file_path):
            with open(file_path, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/octet-stream')
                #print(f'aaaa{os.path.basename(file_path)}')
                response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
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
#     return render(request, 'projeto_teste.html')