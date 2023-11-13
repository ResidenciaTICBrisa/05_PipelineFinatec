import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment,NamedStyle,Border, Side
import os
def pegar_caminho(nome_arquivo):

    # Obter o caminho absoluto do arquivo Python em execução
    caminho_script = os.path.abspath(__file__)

    # Obter o diretório da pasta onde o script está localizado
    pasta_script = os.path.dirname(caminho_script)

    # Combinar o caminho da pasta com o nome do arquivo Excel
    caminho = os.path.join(pasta_script, nome_arquivo)

    return caminho
def estilo_fundep(tabela,tamanho):
    
    # caminho = pegar_caminho(tabela)
    workbook = openpyxl.load_workbook(tabela)
    worksheet = workbook['Relação de despesas']
    size = tamanho + 5

# #Cabecario

#     worksheet.row_dimensions['A1'].height = 19
#     worksheet.row_dimensions['A2'].height = 1
#     worksheet.merge_cells('A1:J1')
#     worksheet['A1'] = 'Rota 2030 - Fundep - Anexo III – Relação de Despesas'

#     for i in range(3,6):
#         worksheet.merge_cells(start_row=i,end_row=i,start_column=1,end_column=2)
    
#     worksheet['A3'] = '1 - Gestora'
#     worksheet['A4'] = '2 - Título do Projeto'
#     worksheet['A5'] = '3 - Coordenador'
#     worksheet.merge_cells('A3:D3')

#Corpo
    count = 1
    for rows in worksheet.iter_rows(min_row=7 ,max_row=size,min_col=2,max_col=2):
        for cell in rows:
            cell.value = count
        count = count + 1
    
    for rows in worksheet.iter_rows(min_row=7 ,max_row=size,min_col=1,max_col=1):
        for cell in rows:
            cell.value = 1
       
    #total despesas nesta
    total_despesa_string = f''
    worksheet.merge_cells()

