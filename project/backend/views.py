from django.shortcuts import render
from .consultas_oracledb import getlimitedRows,getallRows
from django.core.paginator import Paginator,EmptyPage, PageNotAnInteger
# Create your views here.

def project_views(request):
    length = getallRows()
    data = getlimitedRows(length)

    relevant_data = []
    for key, inner_dict in data.items():
        relevant_info = {
        'CODIGO': inner_dict.get('CODIGO', ''),
        'NOME': inner_dict.get('NOME', ''),
        'SALDO': inner_dict.get('SALDO', ''),
        'DATA_ASSINATURA': inner_dict.get('DATA_ASSINATURA', ''),
        'DATA_VIGENCIA': inner_dict.get('DATA_VIGENCIA', ''),
        'DATA_ENCERRAMENTO': inner_dict.get('DATA_ENCERRAMENTO', ''),
        'TIPO_CONTRATO': inner_dict.get('TIPO_CONTRATO', ''),
        'INSTITUICAO_EXECUTORA': inner_dict.get('INSTITUICAO_EXECUTORA', ''),
        'PROCESSO': inner_dict.get('PROCESSO', ''),
        'SUBPROCESSO': inner_dict.get('SUBPROCESSO', ''),
        'COD_PROPOSTA': inner_dict.get('COD_PROPOSTA', ''),
        'PROPOSTA': inner_dict.get('PROPOSTA', ''),
        #'OBJETIVOS': inner_dict.get('OBJETIVOS', ''),
        'VALOR_APROVADO': inner_dict.get('VALOR_APROVADO', ''),
        'NOME_TP_CONTROLE_SALDO': inner_dict.get('NOME_TP_CONTROLE_SALDO', ''),
        'GRUPO_GESTORES': inner_dict.get('GRUPO_GESTORES', ''),
        'GESTOR_RESP': inner_dict.get('GESTOR_RESP', ''),
        'COORDENADOR': inner_dict.get('COORDENADOR', ''),
        'PROCEDIMENTO_COMPRA': inner_dict.get('PROCEDIMENTO_COMPRA', ''),
        'TAB_FRETE': inner_dict.get('TAB_FRETE', ''),
        'TAB_DIARIAS': inner_dict.get('TAB_DIARIAS', ''),
        'CUSTO_OP': inner_dict.get('CUSTO_OP', ''),
        'NOME_FINANCIADOR': inner_dict.get('NOME_FINANCIADOR', ''),
        'DEPARTAMENTO': inner_dict.get('DEPARTAMENTO', ''),
        'SITUACAO': inner_dict.get('SITUACAO', ''),
        'BANCO': inner_dict.get('BANCO', ''),
        'AGENCIA_BANCARIA': inner_dict.get('AGENCIA_BANCARIA', ''),
        'CONTA_BANCARIA': inner_dict.get('CONTA_BANCARIA', ''),
        'CENTRO_CUSTO': inner_dict.get('CENTRO_CUSTO', ''),
        'CONTA_CAIXA': inner_dict.get('CONTA_CAIXA', ''),
        'CATEGORIA_PROJETO': inner_dict.get('CATEGORIA_PROJETO', ''),
        'COD_CONVENIO_CONTA': inner_dict.get('COD_CONVENIO_CONTA', ''),
        'COD_STATUS': inner_dict.get('COD_STATUS', ''),
        'IND_SUB_PROJETO': inner_dict.get('IND_SUB_PROJETO', ''),
        'TIPO_CUSTO_OP': inner_dict.get('TIPO_CUSTO_OP', ''),
        'PROJETO_MAE': inner_dict.get('PROJETO_MAE', ''),
        'ID_COORDENADOR': inner_dict.get('ID_COORDENADOR', ''),
        'ID_FINANCIADOR': inner_dict.get('ID_FINANCIADOR', ''),
        'ID_INSTITUICAO': inner_dict.get('ID_INSTITUICAO', ''),
        'ID_DEPARTAMENTO': inner_dict.get('ID_DEPARTAMENTO', ''),
        'NOME_INSTITUICAO': inner_dict.get('NOME_INSTITUICAO', ''),
        'ID_INSTITUICAO_EXECUTORA': inner_dict.get('ID_INSTITUICAO_EXECUTORA', ''),
        'ID_TIPO': inner_dict.get('ID_TIPO', ''),
    }
        relevant_data.append(relevant_info)
         # Number of items to display per page
    
    search_query = request.GET.get('search', '')  
    if search_query:
        search_results = []
        for entry in relevant_data:
            for key, value in entry.items():
                if key == 'CODIGO' and str(search_query) == str(value):
                    search_results.append(entry)
                    break  
                elif key != 'CODIGO' and isinstance(value, str) and search_query.lower() in value.lower():
                    search_results.append(entry)
                    break  
    else:
        search_results = relevant_data

    items_per_page = 10
    paginator = Paginator(search_results, items_per_page)

    page = request.GET.get('page')
    try:
        items = paginator.get_page(page)
    except PageNotAnInteger:
        items = paginator.get_page(1)
    except EmptyPage:
        items = paginator.get_page(paginator.num_pages)

    return render(request, "backend/projetos.html", {"search_query": search_query,
                                                         "data":items})
