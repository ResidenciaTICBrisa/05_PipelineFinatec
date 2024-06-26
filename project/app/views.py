import os
import shutil
import pyodbc
import datetime
import base64
import io
import re
import tempfile
import json
import os
from PyPDF2 import PdfMerger
import zipfile
from django.conf import settings
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
from .models import Template, Employee
# from .oracle_cruds import consultaPorID
from .new_dev import preenche_planilha,extrair,pegar_caminho
from .preencherFinep import preencheFinep
from .preencheFundep import preenche_fundep
from .preencheFap import preencheFap
from .preencheFub import consultaID,preencheFub,split_archive_name
#from .preencherFinep import preencheFinep
from .capaFub import inserir_round_retangulo
from .capaGeral import inserir_round_retanguloGeral
from django.contrib.admin.models import LogEntry
from .models import UserActivity
from django.core.paginator import Paginator
from django.contrib import messages
from .testeDelete import deletar_arquivos_em_pasta
# from backend.consultas_oracledb import getlimitedRows,getallRows
from backend.consultaSQLServer import consultaCodConvenio, consultaTudo
import pandas as pd
from django.urls import reverse
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from .recibosAutomatizados import acharRecibo, acharReciboDevagar
def pegar_caminho(subdiretorio):
    # Obtém o caminho do script atual
    arq_atual = os.path.abspath(__file__)
    
    # Obtém o diretório do script
    app = os.path.dirname(arq_atual)
    
    # Obtém o diretório pai do script
    project = os.path.dirname(app)
    
    # Obtém o diretório pai do projeto
    pipeline = os.path.dirname(project)
    
    # Junta o diretório pai do projeto com o subdiretório desejado
    caminho_pipeline = os.path.join(pipeline, subdiretorio)
    
    return caminho_pipeline

def pegar_pass(chave):
    arq_atual = os.path.abspath(__file__)
    app = os.path.dirname(arq_atual)
    project = os.path.dirname(app)
    pipeline = os.path.dirname(project)
    desktop = os.path.dirname(pipeline)
    caminho_pipeline = os.path.join(desktop, chave)
    
    return caminho_pipeline



def log_user_activity(user_id, tag, activity):
    UserActivity.objects.create(user_id=user_id, tag=tag, activity=activity)

def convert_datetime_to_string(value):
    if isinstance(value, datetime.datetime):
        return value.strftime('%d/%m/%Y')
    return value
def convert_datetime_to_string2(value):
       # Convert string to datetime object
    date_object = datetime.datetime.strptime(value, "%Y-%m-%d")
    
    # Format the datetime object to the desired format
    formatted_date = date_object.strftime("%d.%m.%Y")
    
    return formatted_date

def extract_strings(input_string):
    # Use regular expressions to find the text before and after '@@'
    matches = re.findall(r'(.*?)@@(.*?)@@', input_string)

    if matches:
        return tuple(matches[0])
    else:
        return (input_string, '')

class HomeView(TemplateView):
    template_name = 'home.html'

@login_required(login_url="/login/")
def user_profile(request):
    if request.method == 'POST':
        user = User.objects.get(username__exact=request.user)
        
        if request.POST.get('new_password1') == request.POST.get('new_password2') and authenticate(username=request.user, password=request.POST.get('old_password')):
            user.set_password(request.POST.get('new_password1'))
            user.save()
            user_authenticated = authenticate(username=request.user, password=request.POST.get('old_password'))
            log_message = f"Alterou a senha"
            log_user_activity(request.user, "Sistema", log_message)
            messages.success(request, "Senha alterada com sucesso!")
            login_a(request, user_authenticated)
        else:
            messages.error(request, "A senha atual inserida, não corresponde com a senha atual. Tente novamente, ou procure o suporte.")

    cpf = Employee.objects.get(user=request.user).cpf
    maskered_cpf = f"{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"

    return render(request,'user_profile.html',{
            "cpf":maskered_cpf,
        })

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
            
            # carrega o dataframe para a sessao
            df = consultaTudo()
            request.session['df'] = df.to_json()
            request.session['codigos'] = consultaCodConvenio(df)
            
            return HttpResponseRedirect ('/projeto/')
        else:
            error_message = 'Usuário ou senha inválido.'

            log_message = f"Tentativa de acesso"
            log_user_activity(usuario, "Sistema", log_message)

            return render(request, 'login.html', {'error_message': error_message})

@login_required(login_url="/login/")
def projeto(request):
    if request.method == 'POST':
        return projeto_legacy(request)
    else:
        # projects = [projeto['CODIGO'] for projeto in relevant_data]  # Obtendo apenas os códigos dos projetos
        return render(request, 'projeto.html', {
            "templates": Template.objects.all(),
            "codigos": request.session['codigos']
        })

def projeto_legacy(request):
    lista_append_db_sql = []
    result = {}
    current_key = None
    mapeamento = None

    codigo = request.POST.get('codigo') # alterar o nome
    template_id = request.POST.get('template')
    consultaInicio = request.POST.get('inicio')
    consultaFim = request.POST.get('fim')

    print(f"----------------------\n{codigo}\n{template_id}\n{consultaInicio}\n{consultaFim}\n{request.POST}\n-----------------------")

    print(type(consultaInicio))
    print(consultaFim)
    emptyText = ""

    try:
        db_fin = consultaID(codigo)
    except:
        print("Erro na consulta codigo inexistente ou invalido")
        # return render(request,'projeto.html',{
        #     "templates":Template.objects.all(),
        # })

    # nome = Template.objects.get(pk=template_id)
    # nome = Template.objects.get(pk=template_id)
    try:
        nome = Template.objects.get(pk=template_id)
    except:
        print("deu ruim")
        # return render(request,'projeto.html',{
        #     "templates":Template.objects.all(),
        # })

    dict_final = {}

    caminho_pasta_planilhas = "../../planilhas/"
    caminhoPastaPlanilhasPreenchidas = "../../planilhas_preenchidas/"

    # Obtém o diretório atual do script
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    #deletar planilhas preendhidas
    testeCaminho = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas)
    deletar_arquivos_em_pasta(testeCaminho)


    #corrigindo datas
    consultaInicial = convert_datetime_to_string2(consultaInicio)
    consultaFinal = convert_datetime_to_string2(consultaFim)
    
    print(consultaFinal)

    if nome.nome_template == "fundep":
    # Combina o diretório atual com o caminho para a pasta "planilhas_preenchidas" e o nome do arquivo
        testeCaminhoFundep = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"ModeloFUNDEP.xlsx")
        preenche_planilha(testeCaminhoFundep,dict_final,codigo,template_id,consultaInicial,consultaFinal,stringNomeFinanciador='FUNDEP')
    if nome.nome_template == "fub":
        testeCaminhoFub = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"Modelo_Fub.xlsx")
        preenche_planilha(testeCaminhoFub,dict_final,codigo,template_id,consultaInicial,consultaFinal,stringNomeFinanciador='FUB')
    if nome.nome_template == "fap":
        testeCaminhoFap = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"modeloFap.xlsx")
        preenche_planilha(testeCaminhoFap,dict_final,codigo,template_id,consultaInicial,consultaFinal,stringNomeFinanciador='FAP')
        
    if nome.nome_template == "finep":
        testeCaminhoFap = os.path.join(diretorio_atual, caminho_pasta_planilhas, f"modeloFinep.xlsx")
        preenche_planilha(testeCaminhoFap,dict_final,codigo,template_id,consultaInicial,consultaFinal,stringNomeFinanciador='FINEP')


    file_path = None
    print(f"download{template_id}")
    if template_id == '1':

        # keys = ['NomeFavorecido','FavorecidoCPFCNPJ','NomeTipoLancamento',
        #         'HisLancamento','NumDocPago','DataEmissao','NumChequeDeposito',
        #         'DataPagamento', 'ValorPago']
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"PC - FUB - {codigo} - {consultaInicial} a {consultaFinal}.xlsx")
        preencheFub(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),file_path)
        inserir_round_retangulo(file_path,consultaInicio,consultaFim,db_fin)

    elif template_id == '2':
        keys = ['NomeFavorecido','FavorecidoCPFCNPJ','NomeRubrica','NumDocPago',
                'DataEmissao','NumChequeDeposito','DataPagamento', 'ValorPago']
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas,f"PC - FUNDEP - {codigo} - {consultaInicial} a {consultaFinal}.xlsx")
        #file_path = pegar_caminho('/home/ubuntu/Desktop/05_PipelineFinatec/planilhas_preenchidas/planilhaPreenchidaModeloFUNDEP.xlsx')
        preenche_fundep(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),keys,file_path)

    elif template_id == '5':
        
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"PC - FINEP - {codigo} - {consultaInicial} a {consultaFinal}.xlsx")
        #file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"PC - FAP - {codigo} - {consultaInicial} a {consultaFinal}.xlsx")
        preencheFinep(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),file_path)
        inserir_round_retanguloGeral(file_path,consultaInicio,consultaFim,db_fin)

    elif template_id == '4':
        
        file_path = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas, f"PC - FAP - {codigo} - {consultaInicial} a {consultaFinal}.xlsx")
        preencheFap(codigo,convert_datetime_to_string(consultaInicio),convert_datetime_to_string(consultaFim),file_path)
        inserir_round_retanguloGeral(file_path,consultaInicio,consultaFim,db_fin)

    else:
        # Handle cases where 'download' doesn't match any expected values
        return HttpResponse("Invalid download request", status=400)

    # Check if the file exists
    if os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/octet-stream')
            response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'

            # adicionando log de consulta
            consulta_log = f"Projeto: {codigo} | Modelo: {nome.nome_template} | Inicio da Prest.: {consultaInicio} | Fim da Prest.: {consultaFim}"
            # LogEntry.objects.log_action(user_id=request.user.id, content_type_id=1, object_repr=consulta_log, action_flag=1, change_message="Consulta de prestação de contas")

            log_user_activity(request.user, "Consulta",consulta_log)

            return response
    else:
        print("Invalid aaaaaaaaaaa request")

    # return render(request,'projeto.html',{
    #     "templates":Template.objects.all(),
    # })

    #delete the files
    
  


def custom_logout(request):
    logout(request)
    return redirect('/')

from .models import UserActivity

def is_admin(user):
    return user.is_authenticated and user.is_staff

@user_passes_test(is_admin)
def user_activity_logs(request):
    logs_list = UserActivity.objects.all()

    # Filter by user_id
    user_id_filter = request.GET.get('user_id')
    if user_id_filter:
        logs_list = logs_list.filter(user_id=user_id_filter)

    # Filter by date
    date_filter = request.GET.get('date')
    if date_filter:
        logs_list = logs_list.filter(timestamp__date=date_filter)

    logs_list = logs_list.order_by('-timestamp')  # Order logs by timestamp in descending order

    paginator = Paginator(logs_list, 50)  # Show 50 logs per page

    page = request.GET.get('page')
    logs = paginator.get_page(page)

    print(type(logs))

    return render(request, 'user_activity_logs.html', {'logs': logs})





def consultaNotas(request, filename):
    '''
    CARREGA PAGINA DAS NOTAS FISCAIS
    '''
    try:
        file_path = pegar_pass("passs.txt")
        conStr = ''
        with open(file_path, 'r') as file:
            conStr = file.readline().strip()
        conn = pyodbc.connect(conStr)
        cursor = conn.cursor()
        
        # queryConsult = f"""
        #     SELECT [Pedido].[CodPedido],
        #            [Pedido].[NumPedido],
        #            [ArquivoBinario].[ArquivoBinario],
        #            [Arquivo].*
        #     FROM [Conveniar].[dbo].[Pedido]
        #     LEFT JOIN [Conveniar].[dbo].[Arquivo]
        #         ON [Pedido].[CodPedido] = [Arquivo].[CodSolicitacao]
        #     LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia]
        #         ON [Arquivo].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
        #     LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario]
        #         ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]
        #     WHERE NumPedido = '{filename}'
        # """
        queryConsult = f"""
             SELECT [Pedido].[CodPedido],
                   [Pedido].[NumPedido],
                   [ArquivoBinario].[ArquivoBinario],
                   [Arquivo].[NomeArquivo] 
            FROM [Conveniar].[dbo].[Pedido]
            LEFT JOIN [Conveniar].[dbo].[Arquivo]
                ON [Pedido].[CodPedido] = [Arquivo].[CodSolicitacao]
            LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia]
                ON [Arquivo].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
            LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario]
                ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]
            WHERE NumPedido = '{filename}'

UNION ALL

SELECT [OPCompraAF].[CodOPCompraAF],
                   [OPCompraAF].[NumOPCompraAF],
                   [ArquivoBinario].[ArquivoBinario],
                   [ArquivoOpCompraAF].[NomeArquivoOpCompraAF]
            FROM [Conveniar].[dbo].[OPCompraAF]
            LEFT JOIN [Conveniar].[dbo].[ArquivoOpCompraAF]
                ON [OPCompraAF].[CodOPCompraAF] = [ArquivoOpCompraAF].[CodOPCompraAF]
            LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia]
                ON [ArquivoOpCompraAF].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
            LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario]
                ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]
            WHERE [OPCompraAF].[NumOPCompraAF] = '{filename}'


        """
        
        cursor.execute(queryConsult)
      
        items = cursor.fetchall()

        #print(f'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{items}')
        cwd = os.getcwd()
        folder_name = "diretoriopdf"
        folder_path = os.path.join(cwd, folder_name)

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Folder '{folder_name}' created successfully!")
        else:
             print(f"Folder '{folder_name}' already exists.")    

        # Merge individual PDFs into one
        merger = PdfMerger()
        count = 0  

        for item in items:
            pdf_data = item[2]
            #print(f"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{pdf_data}")
            if not pdf_data:
                return render(request, 'report.html')
            filename = item[1]

            # Write PDF data to temporary file
            pdf_path = os.path.join(folder_name, f'{filename}{count}.pdf')
            print(pdf_path)
            with open(pdf_path, 'wb') as pdf_file:
                pdf_file.write(pdf_data)

            # Add PDF to merger
            merger.append(pdf_path)
            count += 1  # Increment count


        merged_pdf_filename = 'merged.pdf'
        merged_pdf_path = os.path.join(folder_path, merged_pdf_filename)
        print(merged_pdf_path)
        merger.write(merged_pdf_path)
        merger.close()

        with open(merged_pdf_path, 'rb') as merged_pdf_file:
            response = HttpResponse(merged_pdf_file.read(), content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="nota-{filename}.pdf"'

       
        if os.path.exists(folder_path):
                # Delete the folder and its contents
                shutil.rmtree(folder_path)
                print(f"The folder '{folder_name}' has been successfully deleted.")
        else:
                print(f"The folder '{folder_name}' does not exist.")
                
        return response
    except Exception as e:
        # Handle any exceptions
        #return HttpResponse(f"An error occurred: {str(e)} NOTA NÃO EXISTE")
        return render(request, 'report.html')


# def download_base64(request,filename):
#     if request.method == 'GET':
#         return download_base64_pdf(request,filename)
#     else:
#         # projects = [projeto['CODIGO'] for projeto in relevant_data]  # Obtendo apenas os códigos dos projetos
#         return render(request, 'recibo.html')    
    
def download_base64_pdf(request, filename):
    '''
    CARREGA PAGINA DOS RECIBOS
    
    '''
    try:
   
        
        nota = acharRecibo("selenabot","finatec@300424",filename)
                #print(nota)
        string_with_substring = nota
        substring_to_remove = "data:application/pdf;base64,"
               # Remove the substring
        result_string = string_with_substring.replace(substring_to_remove, "")

                 #print(result_string)


                # Decode the Base64 content
        pdf_content = base64.b64decode(result_string)

                # Create an HTTP response with the PDF content
        response = HttpResponse(pdf_content, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="recibo-{filename}.pdf"'
        return response
    
    except Exception as e:
        # Handle any exceptions
        #return HttpResponse(f"An error occurred: {str(e)} NOTA NÃO EXISTE")
        return render(request, 'recibo.html')

# def download_todos_arquivos(request,filename):
#     '''
#     CARREGA PAGINA DAS NOTAS FISCAIS
#     '''
#     try:
#         file_path = pegar_pass("passs.txt")
#         conStr = ''
#         with open(file_path, 'r') as file:
#             conStr = file.readline().strip()
#         conn = pyodbc.connect(conStr)
#         cursor = conn.cursor()
#         name, date1, date2 = split_archive_name(filename)
    
#         listaPedidos = f"""
#         SELECT
#             NumDocFinConvenio
                
#                 FROM [Conveniar].[dbo].[LisLancamentoConvenio]
#             WHERE [LisLancamentoConvenio].CodConvenio = {name} AND [LisLancamentoConvenio].CodStatus = 27
#             AND [LisLancamentoConvenio].DataPagamento BETWEEN '{date1}' AND '{date2}' and [LisLancamentoConvenio].CodRubrica not in (2,0) 
#             order by DataPagamento

#         """
#         cursor.execute(listaPedidos)
        
#         #print(listaPedidos)

#         NumDocFinConvenios = cursor.fetchall()
#         merged_pdf_responses = []
#         #print(NumDocFinConvenios)
#         cwd = os.getcwd()
#         folder_name = "temp_pdfs"
#         folder_path = os.path.join(cwd, folder_name)

#         if not os.path.exists(folder_path):
#             os.makedirs(folder_path)
#             print(f"Folder '{folder_name}' created successfully!")
#         else:
#              print(f"Folder '{folder_name}' already exists.") 

#         for Doc in NumDocFinConvenios:
#             print(type(Doc))
#             row_str = str(Doc[0])
#             print(row_str)
#             queryConsult = f"""
#              SELECT [Pedido].[CodPedido],
#                    [Pedido].[NumPedido],
#                    [ArquivoBinario].[ArquivoBinario],
#                    [Arquivo].[NomeArquivo] 
#             FROM [Conveniar].[dbo].[Pedido]
#             LEFT JOIN [Conveniar].[dbo].[Arquivo]
#                 ON [Pedido].[CodPedido] = [Arquivo].[CodSolicitacao]
#             LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia]
#                 ON [Arquivo].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
#             LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario]
#                 ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]
#             WHERE NumPedido = '{row_str}'

#             UNION ALL

#             SELECT [OPCompraAF].[CodOPCompraAF],
#                    [OPCompraAF].[NumOPCompraAF],
#                    [ArquivoBinario].[ArquivoBinario],
#                    [ArquivoOpCompraAF].[NomeArquivoOpCompraAF]
#             FROM [Conveniar].[dbo].[OPCompraAF]
#             LEFT JOIN [Conveniar].[dbo].[ArquivoOpCompraAF]
#                 ON [OPCompraAF].[CodOPCompraAF] = [ArquivoOpCompraAF].[CodOPCompraAF]
#             LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia]
#                 ON [ArquivoOpCompraAF].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
#             LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario]
#                 ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]
#             WHERE [OPCompraAF].[NumOPCompraAF] = '{row_str}'


#         """ 
#             #print(queryConsult)
#             cursor.execute(queryConsult)
#             items = cursor.fetchall()
#             #print(items)

            
         
#             pdf_files = []
#               # Merge individual PDFs into one
#             merger = PdfMerger()
#             count = 0  
#             conta = 0
#             for item in items:
#                 filename = item[1]
#                 pdf_data = item[2]
#                 #print(f"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{pdf_data}")
#                 if not pdf_data:
                     
#                     nota = acharRecibo("selenabot","finatec@300424",filename)
#                             #print(nota)
#                     string_with_substring = nota
#                     substring_to_remove = "data:application/pdf;base64,"
#                         # Remove the substring
#                     result_string = string_with_substring.replace(substring_to_remove, "")

#                     pdf_content = base64.b64decode(result_string)
#                     pdf_path = os.path.join(folder_name, f'{filename}{conta}.pdf')
#                     print(pdf_path)
#                     with open(pdf_path, 'wb') as pdf_file:
#                         pdf_file.write(pdf_content)

                   
#                 else:
#                     # Write PDF data to temporary file
#                     pdf_path = os.path.join(folder_name, f'{filename}{count}.pdf')
#                     print(pdf_path)
#                     with open(pdf_path, 'wb') as pdf_file:
#                         pdf_file.write(pdf_data)

#                     # Add PDF to merger
                    
        
#             # Create a temporary file path for the zip file
#             zip_file_path = os.path.join(cwd, 'temp_pdfs.zip')

#             # Create a zip file
#             with zipfile.ZipFile(zip_file_path, 'w') as zipf:
#                 for root, dirs, files in os.walk(folder_path):
#                     for file in files:
#                         file_path = os.path.join(root, file)
#                         zipf.write(file_path, os.path.relpath(file_path, folder_path))

#             # Open the zip file
#             with open(zip_file_path, 'rb') as f:
#                 response = HttpResponse(f.read(), content_type='application/zip')
#                 response['Content-Disposition'] = f'attachment; filename={name}_{date1}_{date2}.zip'

#             # Clean up: remove the temporary zip file
#             os.remove(zip_file_path)
#             if os.path.exists(folder_path):
#                 # Delete the folder and its contents
#                 shutil.rmtree(folder_path)
#                 print(f"The folder '{folder_name}' has been successfully deleted.")
#             else:
#                 print(f"The folder '{folder_name}' does not exist.")
#             return response
        
        
#     except Exception as e:
        # Handle any exceptions
        return HttpResponse(f"An error occurred: {str(e)}")