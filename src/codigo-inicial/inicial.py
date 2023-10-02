import os
import openpyxl as op

# Adicionando o nome da gestora do projeto para FINATEC
# Ajustando caminho --------------------------------------

    # Obter o caminho absoluto do arquivo Python em execução
caminho_script = os.path.abspath(__file__)

    # Obter o diretório da pasta onde o script está localizado
pasta_script = os.path.dirname(caminho_script)

    # Nome do arquivo Excel
nome_arquivo_excel = 'FUNDEP.xlsx'

    # Combinar o caminho da pasta com o nome do arquivo Excel
caminho_arquivo_excel = os.path.join(pasta_script, nome_arquivo_excel)

    # Carregar o arquivo Excel




# Código inicial------------------------------------------

workbook = op.load_workbook(caminho_arquivo_excel)

sheet = workbook.active

gestora = sheet['C3']
gestora.value = "Finatec"


nome_projeto = sheet['C4']
nome_projeto.value = "Projeto X"

coordenador = sheet['C5']
coordenador.value = "Suellen"

referencia = sheet['F3']
referencia.value = "27193*14"

# DÚVIDA:
# Nº Acordo de Parceria 
# Nº Acordo	

# Salvando arquivo na mesma pasta
nome_novo_excel = 'FUNDEP-inicial.xlsx'
caminho_novo_excel = os.path.join(pasta_script, nome_novo_excel)
workbook.save(caminho_novo_excel)