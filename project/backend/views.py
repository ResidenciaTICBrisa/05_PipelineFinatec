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
        'NOME_FINANCIADOR': inner_dict.get('NOME_FINANCIADOR', ''),
        'DATA_ASSINATURA': inner_dict.get('DATA_ASSINATURA', ''),
        'DATA_VIGENCIA': inner_dict.get('DATA_VIGENCIA', ''),
        'COORDENADOR': inner_dict.get('COORDENADOR', ''),
        'VALOR_APROVADO': inner_dict.get('VALOR_APROVADO', ''),
        'GRUPO_GESTORES': inner_dict.get('GRUPO_GESTORES', ''),
    
        
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
