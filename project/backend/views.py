from django.shortcuts import render
from .consultas_oracledb import getlimitedRows,getallRows
from django.core.paginator import Paginator,EmptyPage, PageNotAnInteger
from .consultaSQLServer import consultaLimitedDict
import pandas as pd
# Create your views here.

def project_views(request):
    length = pd.read_json(request.session['df']).shape[0]
    data = pd.read_json(request.session['df'])
    # data = consultaLimitedDict(pd.read_json(request.session['df']), length)
    print(pd.read_json(request.session['df']))

    relevant_data = []
    for index, row in data.iterrows():
        relevant_info = {
            'CODIGO': row['CodConvenio'],
            'NOME': row['NomeConvenio'],
            'NOME_FINANCIADOR': row['NomePessoaFinanciador'],
            'DATA_ASSINATURA': row['DataAssinatura'],
            'DATA_VIGENCIA': row['DataVigencia'],
            'COORDENADOR': row['NomePessoaResponsavel'],
            # 'VALOR_APROVADO': row['VALOR_APROVADO'],
            'GRUPO_GESTORES': row['NomeGrupoGestor'],
        }
        relevant_data.append(relevant_info)
         # Number of items to display per page
    print(relevant_data)
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
