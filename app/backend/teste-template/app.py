import os
import openpyxl as op

# Adicionando o nome da gestora do projeto para FINATEC

# -------------------------------------------------------------

# workbook = op.load_workbook('teste-template/FUNDEP.xlsx')

# sheet = workbook.active

# gestora = sheet['C3']
# gestora.value = "Finatec"


# nome_projeto = sheet['C4']
# nome_projeto.value = "Projeto X"

# coordenador = sheet['C5']
# coordenador.value = "Suellen"

# referencia = sheet['F3']
# referencia.value = "Suellen"

# DÚVIDA:
# Nº Acordo de Parceria 
# Nº Acordo	

# print(cell.value)

# workbook.save('FUNDEP-preenchido.xlsx')

# -------------------------------------------------------------

def pegar_caminho(nome_arquivo):

    # Obter o caminho absoluto do arquivo Python em execução
    caminho_script = os.path.abspath(__file__)

    # Obter o diretório da pasta onde o script está localizado
    pasta_script = os.path.dirname(caminho_script)

    # Combinar o caminho da pasta com o nome do arquivo Excel
    caminho = os.path.join(pasta_script, nome_arquivo)

    return caminho

def preenche_cell(planilha, celula, text):

    caminho = pegar_caminho(planilha)

    # carrega a planilha de acordo com o caminho
    workbook = op.load_workbook(caminho)

    sheet = workbook.active

    cell = sheet[celula]
    cell.value = f'{text}'

    planilha_preenchida = pegar_caminho('preenchido-'+planilha)

    workbook.save(planilha_preenchida)
    


preenche_cell('FUNDEP.xlsx', 'C3', 'FUNDEP')


    


