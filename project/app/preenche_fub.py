import pyodbc
from datetime import datetime,date
import openpyxl
import os
from .estilo_fub import *
from collections import defaultdict
from .oracle_cruds import getAnalistaDoProjetoECpfCoordenador #nao ta usando
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL


def formatar_cpf(cpf):
    cpf_formatado = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
    return cpf_formatado
def check_format(time_data, format='%Y-%m-%d'):
    try:
        # Try to parse the time_data using the specified format
        datetime.strptime(time_data, format)
        return True  # The time_data matches the format
    except ValueError:
        return False  # The time_data does not match the format

# def pegar_caminho(nome_arquivo):
 
#     # Obter o caminho absoluto do arquivo Python em execução
#     caminho_script = os.path.abspath(__file__)

#     # Obter o diretório da pasta onde o script está localizado
#     pasta_script = os.path.dirname(caminho_script)

#     # Combinar o caminho da pasta com o nome do arquivo Excel
#     caminho = os.path.join(pasta_script, nome_arquivo)

#     return caminho

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

def convert_datetime_to_string(value):
    if isinstance(value, datetime):
        return value.strftime('%d/%m/%Y')
    return value

def pegar_pass(chave):
    arq_atual = os.path.abspath(__file__)
    app = os.path.dirname(arq_atual)
    project = os.path.dirname(app)
    pipeline = os.path.dirname(project)
    desktop = os.path.dirname(pipeline)
    caminho_pipeline = os.path.join(desktop, chave)
    
    return caminho_pipeline

def consultaConciliacaoBancaria(IDPROJETO, DATA1, DATA2):
    ''' Função que vai pega os dados da Rubrica 9 Despesas Financeiras e transformalos em dataframe
    para poder popular a databela Despesas Financeiras
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
    '''
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    consultaSemEstorno = f"SELECT DISTINCT DataPagamento,ValorLancamento,NumDocFinConvenio,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) NOT LIKE '%estorno%' order by DataPagamento"
    consultaComEstorno =  f"SELECT DISTINCT DataPagamento,ValorLancamento,NumDocFinConvenio,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) LIKE '%estorno%' order by DataPagamento"
    dfSemEstorno = pd.read_sql(consultaSemEstorno, engine, params=parametros)
    dfComEstorno = pd.read_sql(consultaComEstorno, engine, params=parametros)
   

    return dfSemEstorno,dfComEstorno

def consultaLisLancamentoConveniarSemData():

    #file_path = "/home/ubuntu/Desktop/devfront/devfull/pass.txt"
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    # print(conStr)
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    # print(cursor)

    # SQL query
    
    #sql = f"SELECT TOP 1 * FROM [Conveniar].[dbo].[LisConvenio]"
    sql = f"SELECT TOP 1 [LisConvenio].* , [LisPessoa].[CPFCNPJ] as 'CPFCoordenador' FROM [Conveniar].[dbo].[LisConvenio] INNER JOIN  [Conveniar].[dbo].[LisUsuario] ON [LisConvenio].[CodUsuarioResponsavel] = [LisUsuario].[CodUsuario] INNER JOIN  [Conveniar].[dbo].[LisPessoa] ON [LisUsuario].[CodPessoa] = [LisPessoa].[CodPessoa]"

    # Execute the query
    cursor.execute(sql)

    # # Fetch the results
    # result = cursor.fetchall()

    # Close the cursor and connection
    # cursor.close()
    # conn.close()

    return cursor

def consultaLisLancamentoConveniar(IDPROJETO, DATA1, DATA2):

 
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
   

    # SQL querys
    
    # formatted_date1 = DATA1.strftime("%Y-%m-%d")
    # formatted_date2 = DATA2.strftime("%Y-%m-%d")
    sql = f"SELECT * FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ?  ORDER BY CodLancamento"
    #sql = f"SELECT * FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ?  ORDER BY NumDocFinConvenio"

    # Execute the query
    cursor.execute(sql, IDPROJETO, DATA1, DATA2)


    return cursor
    
    # return records

def consultaID(IDPROJETO):

   #file_path = "/home/ubuntu/Desktop/devfront/devfull/pass.txt"
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    
   
    consulta = {}
   

    # SQL querys
    
    sql = f"SELECT [LisConvenio].* , [LisPessoa].[CPFCNPJ] as 'CPFCoordenador' FROM [Conveniar].[dbo].[LisConvenio] INNER JOIN  [Conveniar].[dbo].[LisUsuario] ON [LisConvenio].[CodUsuarioResponsavel] = [LisUsuario].[CodUsuario] INNER JOIN  [Conveniar].[dbo].[LisPessoa] ON [LisUsuario].[CodPessoa] = [LisPessoa].[CodPessoa] WHERE CodConvenio = ? "

    # Execute the query
    cursor.execute(sql, IDPROJETO)


    records = cursor.fetchall()
    
    collums = cursor.description
    

    for i in range(len(collums)):
        consulta[collums[i][0]] = records[0][i]
  
    cursor.close()
    conn.close()
    print("The connection is closed")
    #print(consulta)
    
    # return records
    return consulta
#retorna todos os valores dos dicionarios
def get_values_from_dict(codigo,data1,data2):
  
    gete = consultaLisLancamentoConveniar(codigo,data1,data2)

    collums = []
    for i in gete.description:
        collums.append(i[0])
    
    #print(collums)

    value = []
    for i in gete:
        val = tuple(convert_datetime_to_string(item) for item in i)
        value.append(val)
    #print(value)
    list_of_dicts = [dict(zip(collums, values)) for values in value]

    #print(list_of_dicts)
    return list_of_dicts
#retorna os valores dado uma chave, por exmeplo se for VALOR_PAGO = 4,50
def retornavalores(list_of_dicts,keys):
    values = [d.get(key) for d in list_of_dicts for key in keys]
    
    #print(values)
    return values
#separa  os dics por rubrica, por exemplo caso queira acessar a da rubrica 87 a= separarporrubrica() - > a[87]
def separarporrubrica(codigo,data1,data2):
    valor = get_values_from_dict(codigo,data1,data2)


    #  Extract unique values from the 'ID_RUBRICA' key
    unique_id_rubrica_values = set(item['CodRubrica'] for item in valor)

    # Create separate lists of dictionaries for each unique 'ID_RUBRICA' value
    categorized_data = {value: [] for value in unique_id_rubrica_values}
    for item in valor:
        categorized_data[item['CodRubrica']].append(item)
    
    return categorized_data
#separa por tipo de favorecido as rubricas 87 e 9
def tipodefavorecido(codigo,data1,data2):
    data_categorizada = separarporrubrica(codigo,data1,data2)
    #print(data_categorizada)
    if 87 not in data_categorizada or not data_categorizada[87]:
        print("Data not available or empty.")
        return None  # or handle the case accordingly
    separarportipodefavorecido = set(item['TIPO_FAVORECIDO'] for item in data_categorizada[87])
    #print(separarportipodefavorecido)

    # # Step 2: Create separate lists of dictionaries for each unique 'ID_RUBRICA' value
    dict_favorecido_fisica_e_juridica = {value: [] for value in separarportipodefavorecido}
    for item in data_categorizada[87]:
        dict_favorecido_fisica_e_juridica[item['TIPO_FAVORECIDO']].append(item)

    #print(dict_favorecido_fisica_e_juridica)
    return dict_favorecido_fisica_e_juridica

def queryReceitaXDespesa(CodConvenio,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(ValorPago) AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio, DATA1, DATA2)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor

def queryReceitaXDespesaTotal(CodConvenio,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(ValorPago) AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento <= ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio, DATA2)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor
def queryReceitaXDespesa(CodConvenio,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(ValorPago) AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio, DATA1, DATA2)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor

def queryReceitaXDespesaTotal(CodConvenio,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(ValorPago) AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento <= ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio, DATA2)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor

def queryValorPrevisto(CodConvenio):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(VALOR)AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisConvenioItemAprovado] WHERE CodConvenio = ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor

def resumoReceitaDespesa(rubricaprincipal,rubricas,rubricaprincipalstring,consulta_list):
 
   
    valor = consulta_list

    #  Extract unique values from the 'ID_RUBRICA' key
    unique_id_rubrica_values = set(item['CodRubrica'] for item in valor)

    # Create separate lists of dictionaries for each unique 'ID_RUBRICA' value
    categorized_data = {value: [] for value in unique_id_rubrica_values}
    for item in valor:
        categorized_data[item['CodRubrica']].append(item)
    
    dicionariosaida = {}
    if rubricaprincipal in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[rubricaprincipal].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if rubricaprincipal not in categorized_data:
                    categorized_data[rubricaprincipal] = categorized_data[num]
                else:
                    categorized_data[rubricaprincipal].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
    

    keys = ['VALOR_TOTAL']
    
    soma = 0
    for i in keys:
        li = [i]
        if rubricaprincipal not in categorized_data or not categorized_data[rubricaprincipal]:
                print("Data not available or empty.")
        else:
            valores_preenchimento = retornavalores(categorized_data[rubricaprincipal],li)
            # print(valores_preenchimento)
            for i in range(len(valores_preenchimento)):
                soma = soma + valores_preenchimento[i]
    
    dicionariosaida[rubricaprincipalstring] = soma

  
    return dicionariosaida
     
def preencherCapa(codigo,planilha):
    analista = getAnalistaDoProjetoECpfCoordenador(codigo)
    caminho = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(caminho)
    sheet = workbook['Capa Finatec']
    sheet['E26'] = analista['NOME_ANALISTA']
    workbook.save(planilha)
    workbook.close()
#preenche planilha de referencia pra nome do coordenador e diretor
def criaReceitaXDespesa(planilha,codigo,data1,data2,tamanhoResumo,dicionarioComDadosEntreAsDatas,dicionarioComDadosEntreAsDatasComMaterial):
    
    caminho = pegar_caminho(planilha)
   
    tamanho,tamanhoequipamentos = estiloReceitaXDespesa(caminho,tamanhoResumo)  
    #Plan = planilha
    # carrega a planilha de acordo com o caminho
    workbook = openpyxl.load_workbook(caminho)
    sheet = workbook['Receita x Despesa']
    input_date = []
    output_date_str = []
    input_date2  = []
    output_date_str2 = []
    if check_format(data1):
        input_date = datetime.strptime(data1, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str = input_date.strftime("%d/%m/%Y")
    else :
         return None
    if check_format(data2):
        input_date2 = datetime.strptime(data2, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str2 = input_date2.strftime("%d/%m/%Y")
    else :
         return None
    
    
    
    string_periodo = f"Período que abrange esta prestação: {output_date_str} a {output_date_str2}"
    sheet['A7'] = string_periodo
    consulta_coordenador = consultaID(codigo)
    # print(consulta_coordenador)
    # print(type(consulta_coordenador))
    stringCoordenador= f'H{tamanho+3}' # retorna lugar do coordanor
    stringTamanhoCPF = f'H{tamanho+5}' # retorna lugar do coordanor
    sheet[stringCoordenador] = consulta_coordenador['NomePessoaResponsavel']
    sheet[stringTamanhoCPF] = formatar_cpf(consulta_coordenador['CPFCoordenador'])
    string_titulo = f"Título do Projeto: {consulta_coordenador['NomeConvenio']}"
    string_executora = f"Executora: {consulta_coordenador['NomePessoaFinanciador']}"
    string_participe = f"Partícipe: Fundação de Empreendimentos Científicos e Tecnológicos - FINATEC"
   # Convert 'DataAssinatura' to "dd/mm/YYYY" format
    datetime_obj1 = consulta_coordenador['DataAssinatura']
    formatted_date1 = datetime_obj1.strftime("%d/%m/%Y")

    # Convert 'DataVigencia' to "dd/mm/YYYY" format
    datetime_obj2 = consulta_coordenador['DataVigencia']
    formatted_date2 = datetime_obj2.strftime("%d/%m/%Y")

# Create the string representing the period of execution
    string_periodo = f"Período de Execução Físico-Financeiro: {formatted_date1} a {formatted_date2}"
    sheet['A3'] = string_titulo
    sheet['A4'] = string_executora
    sheet['A5'] = string_participe
    sheet['A6'] = string_periodo
 
    #dadosquefaltam = getAnalistaDoProjetoECpfCoordenador(codigo)
    #print(f'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{dadosquefaltam}')
    #sheet['H47'] = formatar_cpf(dadosquefaltam["CPF_COORDENADOR"])
    meses_dict = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}   
    strintT = tamanho
    stringTamanho = f'A{tamanho}' # retorna lugar de brasilia
    hoje = date.today()
    data_formatada = f"{hoje.day} de {meses_dict[hoje.month]} de {hoje.year}"
    sheet[stringTamanho] = f'Brasilia, {data_formatada}'


    #Preenchendo a planilha
    # print("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEXEC")
    #print(dicionarioComDadosEntreAsDatas)
    rowInicial = 16
    for key,value in dicionarioComDadosEntreAsDatas.items():
        if key == "Encargos - ISS 5% ":
              sheet['C17'] = key
              sheet['E17'] = value
        elif key == "Encargos - ISS 2% ":
              sheet['C18'] = key
              sheet['E18'] = value
        else:
            sheet[f"H{rowInicial}"] = key
            sheet[f'I{rowInicial}'] = value
    
        rowInicial = rowInicial+1
             
    # print(dicionarioComDadosEntreAsDatasComMaterial)
    
    for key,value in dicionarioComDadosEntreAsDatasComMaterial.items():
        if key == "Receitas":
              sheet['C16'] = key
              sheet['E16'] = value
        elif key == 'Equipamento Material Permanente':
              sheet[f'I{tamanhoequipamentos}'] = value
        elif key == 'Material Permanente e Equipamento Nacional':
              sheet[f'I{tamanhoequipamentos+1}'] = value
        elif key == 'Material Permanente e Equipamento Importado':
              sheet[f'I{tamanhoequipamentos+2}'] = value
        elif key == 'Rendimentos de Aplicações Financeiras':
             sheet[f'A{tamanhoequipamentos+6}'] = 'RF Ref DI Plus Ágil '
             sheet[f'E{tamanhoequipamentos+6}'] = value

              

    workbook.save(planilha)
    workbook.close()

    return strintT
    
def execReceitaDespesa(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho,variavelResumo,variavelResumoTotalAteadata,valorPrevisto):
    
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Exec. Receita e Despesa")
    workbook.save(tabela)
    workbook.close()




    ####VALOR PREVISTO
    
    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',valorPrevisto)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',valorPrevisto)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',valorPrevisto)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',valorPrevisto)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',valorPrevisto)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',valorPrevisto)
    #rendimentodeapliaçõesfinanceiras
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',valorPrevisto)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',valorPrevisto)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',valorPrevisto)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',valorPrevisto) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',valorPrevisto) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',valorPrevisto)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',valorPrevisto)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',valorPrevisto)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',valorPrevisto)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',valorPrevisto)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',valorPrevisto)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Importado',valorPrevisto)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',valorPrevisto)

    # print(dict14)

    dicionarioValorPrevisto = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    mergeTotalPrevisto = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19}

    dicionarioValoresPrevisto = {}
    for key, value in dicionarioValorPrevisto.items() :
        if value != 0: 
            dicionarioValoresPrevisto[key]= value 




    ##########PREENCHE COM BETWEEN - REALIZADO EXECTUADO NO PERIODO
    #pessoal fisica
    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',variavelResumo)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',variavelResumo)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',variavelResumo)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',variavelResumo)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',variavelResumo)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',variavelResumo)
    #rendimentodeapliaçõesfinanceiras
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',variavelResumo)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',variavelResumo)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',variavelResumo)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',variavelResumo) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',variavelResumo) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',variavelResumo)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',variavelResumo)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',variavelResumo)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',variavelResumo)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',variavelResumo)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',variavelResumo)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Nacional',variavelResumo)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',variavelResumo)
    #Receitas
    dict20 = resumoReceitaDespesa(2,rubricas,'Receitas',variavelResumo)
    #Despesas financeiras
    dict21 = resumoReceitaDespesa(9,rubricas,'Despesas Financeiras',variavelResumo)



    dicionarioQueVaiSerPrintado = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    merged_dict = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19,**dict20,**dict21}




    dictPraCalcularTamanho = {}
    for key, value in dicionarioQueVaiSerPrintado.items() :
        if value != 0: 
            dictPraCalcularTamanho[key]= value 
    
    newDictValoresPrevisto = dict(dicionarioValoresPrevisto)
    newDictPraCalcularTamanho = dict(dictPraCalcularTamanho)
    
    newDictPraCalcularTamanho.update(dicionarioValoresPrevisto)
    newDictPraCalcularTamanho.update(dictPraCalcularTamanho)

    tamanho = len(newDictPraCalcularTamanho)
    #print(dictPraCalcularTamanho)
    
    stringTamanho = tamanho + 16 
    estiloExecReceitaDespesa(tabela,tamanho,stringTamanho)


    ##############SEM O BETWEEN REALIZADO ACUMULADO ATE O PERIODO

    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',variavelResumoTotalAteadata)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',variavelResumoTotalAteadata)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',variavelResumoTotalAteadata)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',variavelResumoTotalAteadata)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',variavelResumoTotalAteadata)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',variavelResumoTotalAteadata)
    #Rendimento de aplicação
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',variavelResumoTotalAteadata)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',variavelResumoTotalAteadata)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',variavelResumoTotalAteadata)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',variavelResumoTotalAteadata) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',variavelResumoTotalAteadata) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',variavelResumoTotalAteadata)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',variavelResumoTotalAteadata)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',variavelResumoTotalAteadata)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',variavelResumoTotalAteadata)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',variavelResumoTotalAteadata)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',variavelResumoTotalAteadata)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Nacional',variavelResumoTotalAteadata)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',variavelResumoTotalAteadata)



    dicionarioAcumuladoAteOPeriodo = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    mergeTotal = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19}

    dicionarioValoresTotais = {}
    for key, value in dicionarioAcumuladoAteOPeriodo.items() :
        if value != 0: 
            dicionarioValoresTotais[key]= value 
    

    #preencher
    workbook = openpyxl.load_workbook(planilha)
    sheet = workbook['Exec. Receita e Despesa']

    input_date = []
    output_date_str = []
    input_date2  = []
    output_date_str2 = []
    if check_format(data1):
        input_date = datetime.strptime(data1, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str = input_date.strftime("%d/%m/%Y")
    else :
         return None
    if check_format(data2):
        input_date2 = datetime.strptime(data2, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str2 = input_date2.strftime("%d/%m/%Y")
    else :
         return None
    
    
    
    string_periodo = f"Período que abrange esta prestação: {output_date_str} a {output_date_str2}"
    sheet['A7'] = string_periodo
    consulta_coordenador = consultaID(codigo)
    # print(consulta_coordenador)
    # print(type(consulta_coordenador))
    stringCoordenador= f'F{stringTamanho+11}' # retorna lugar do coordanor
    stringCoordanadorCargo = f'F{stringTamanho+12}'
    sheet[stringCoordanadorCargo] = f"Coordenador(a)"
    stringTamanhoCPF = f'F{stringTamanho+13}' # retorna lugar do coordanor
    sheet[stringCoordenador] = consulta_coordenador['NomePessoaResponsavel']
    sheet[stringTamanhoCPF] = formatar_cpf(consulta_coordenador['CPFCoordenador'])
    string_titulo = f"Título do Projeto: {consulta_coordenador['NomeConvenio']}"
    string_executora = f"Executora: {consulta_coordenador['NomePessoaFinanciador']}"
    string_participe = f"Partícipe: Fundação de Empreendimentos Científicos e Tecnológicos - FINATEC"
   # Convert 'DataAssinatura' to "dd/mm/YYYY" format
    datetime_obj1 = consulta_coordenador['DataAssinatura']
    formatted_date1 = datetime_obj1.strftime("%d/%m/%Y")

    # Convert 'DataVigencia' to "dd/mm/YYYY" format
    datetime_obj2 = consulta_coordenador['DataVigencia']
    formatted_date2 = datetime_obj2.strftime("%d/%m/%Y")

# Create the string representing the period of execution
    string_periodo = f"Período de Execução Físico-Financeiro: {formatted_date1} a {formatted_date2}"
    sheet['A3'] = string_titulo
    sheet['A4'] = string_executora
    sheet['A5'] = string_participe
    sheet['A6'] = string_periodo
 
    #dadosquefaltam = getAnalistaDoProjetoECpfCoordenador(codigo)
    #print(f'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{dadosquefaltam}')
    #sheet['H47'] = formatar_cpf(dadosquefaltam["CPF_COORDENADOR"])
    meses_dict = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}   

    stringTamanhoBrasilia = f'A{stringTamanho+10}' # retorna lugar de brasilia
    hoje = date.today()
    data_formatada = f"{hoje.day} de {meses_dict[hoje.month]} de {hoje.year}"
    sheet[stringTamanhoBrasilia] = f'Brasilia, {data_formatada}'

    #despesas correntes printr
    imprimirResumoLinha = 16


    # print("dictpracalcularTamanho")
    # print(dictPraCalcularTamanho)
    # print("dictprevisto")
    # print(dicionarioValoresPrevisto)



    #REALIZADO
    for key,value in newDictPraCalcularTamanho.items():
          if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringChave = f"A{imprimirResumoLinha}"
            imprimirResumoLinha = 1 + imprimirResumoLinha
            sheet[posicaoStringChave]=key


    imprimirResumoLinha = 16

    for key,value in dictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"C{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    posicaoEquipamentoMaterialPermanente = 16 + tamanho + 3
    for key,value in merged_dict.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value

    #REALIZADO ACUMULADO ATE O PERIODO
    imprimirResumoLinha = 16
    for key,value in dicionarioValoresTotais.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"G{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    for key,value in mergeTotal.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value

 
             
    #valor previsto            
    imprimirResumoLinha = 16
    newDictPraCalcularTamanho.update(dicionarioValoresPrevisto)
    for key,value in newDictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"B{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    imprimirResumoLinha = 16
    for key,value in newDictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"F{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha


    for key,value in mergeTotalPrevisto.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value



    workbook.save(planilha)
    workbook.close()
    return tamanho,dictPraCalcularTamanho,merged_dict

def queryValorPrevisto(CodConvenio):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = pyodbc.connect(conStr)
    cursor = conn.cursor()
    consulta = {}

    # SQL querys
    
    sql = f"SELECT SUM(VALOR)AS VALOR_TOTAL, CodRubrica, NomeRubrica FROM [Conveniar].[dbo].[LisConvenioItemAprovado] WHERE CodConvenio = ? GROUP BY NomeRubrica, CodRubrica"
   
    # Execute the query
    queryRXD = cursor.execute(sql, CodConvenio)

    collums = []
    for i in queryRXD.description:
        collums.append(i[0])
    records = cursor.fetchall()

    consulta_list = []

    for i in range(len(records)):
        consulta = {}  # Create a new dictionary for each iteration of the outer loop
        for j in range(3):
            consulta[collums[j]] = records[i][j]
        consulta_list.append(consulta)
   
    valor = consulta_list

    return valor

def resumoReceitaDespesa(rubricaprincipal,rubricas,rubricaprincipalstring,consulta_list):
 
   
    valor = consulta_list

    #  Extract unique values from the 'ID_RUBRICA' key
    unique_id_rubrica_values = set(item['CodRubrica'] for item in valor)

    # Create separate lists of dictionaries for each unique 'ID_RUBRICA' value
    categorized_data = {value: [] for value in unique_id_rubrica_values}
    for item in valor:
        categorized_data[item['CodRubrica']].append(item)
    
    dicionariosaida = {}
    if rubricaprincipal in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[rubricaprincipal].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if rubricaprincipal not in categorized_data:
                    categorized_data[rubricaprincipal] = categorized_data[num]
                else:
                    categorized_data[rubricaprincipal].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
    

    keys = ['VALOR_TOTAL']
    
    soma = 0
    for i in keys:
        li = [i]
        if rubricaprincipal not in categorized_data or not categorized_data[rubricaprincipal]:
                print("Data not available or empty.")
        else:
            valores_preenchimento = retornavalores(categorized_data[rubricaprincipal],li)
            # print(valores_preenchimento)
            for i in range(len(valores_preenchimento)):
                soma = soma + valores_preenchimento[i]
    
    dicionariosaida[rubricaprincipalstring] = soma

  
    return dicionariosaida
     
def preencherCapa(codigo,planilha):
    analista = getAnalistaDoProjetoECpfCoordenador(codigo)
    caminho = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(caminho)
    sheet = workbook['Capa Finatec']
    sheet['E26'] = analista['NOME_ANALISTA']
    workbook.save(planilha)
    workbook.close()
#preenche planilha de referencia pra nome do coordenador e diretor
def criaReceitaXDespesa(planilha,codigo,data1,data2,tamanhoResumo,dicionarioComDadosEntreAsDatas,dicionarioComDadosEntreAsDatasComMaterial):
    
    caminho = pegar_caminho(planilha)
   
    tamanho,tamanhoequipamentos = estiloReceitaXDespesa(caminho,tamanhoResumo)  
    #Plan = planilha
    # carrega a planilha de acordo com o caminho
    workbook = openpyxl.load_workbook(caminho)
    sheet = workbook['Receita x Despesa']
    input_date = []
    output_date_str = []
    input_date2  = []
    output_date_str2 = []
    if check_format(data1):
        input_date = datetime.strptime(data1, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str = input_date.strftime("%d/%m/%Y")
    else :
         return None
    if check_format(data2):
        input_date2 = datetime.strptime(data2, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str2 = input_date2.strftime("%d/%m/%Y")
    else :
         return None
    
    
    
    string_periodo = f"Período que abrange esta prestação: {output_date_str} a {output_date_str2}"
    sheet['A7'] = string_periodo
    consulta_coordenador = consultaID(codigo)
    # print(consulta_coordenador)
    # print(type(consulta_coordenador))
    stringCoordenador= f'H{tamanho+3}' # retorna lugar do coordanor
    stringTamanhoCPF = f'H{tamanho+5}' # retorna lugar do coordanor
    sheet[stringCoordenador] = consulta_coordenador['NomePessoaResponsavel']
    sheet[stringTamanhoCPF] = formatar_cpf(consulta_coordenador['CPFCoordenador'])
    string_titulo = f"Título do Projeto: {consulta_coordenador['NomeConvenio']}"
    string_executora = f"Executora: {consulta_coordenador['NomePessoaFinanciador']}"
    string_participe = f"Partícipe: Fundação de Empreendimentos Científicos e Tecnológicos - FINATEC"
   # Convert 'DataAssinatura' to "dd/mm/YYYY" format
    datetime_obj1 = consulta_coordenador['DataAssinatura']
    formatted_date1 = datetime_obj1.strftime("%d/%m/%Y")

    # Convert 'DataVigencia' to "dd/mm/YYYY" format
    datetime_obj2 = consulta_coordenador['DataVigencia']
    formatted_date2 = datetime_obj2.strftime("%d/%m/%Y")

# Create the string representing the period of execution
    string_periodo = f"Período de Execução Físico-Financeiro: {formatted_date1} a {formatted_date2}"
    sheet['A3'] = string_titulo
    sheet['A4'] = string_executora
    sheet['A5'] = string_participe
    sheet['A6'] = string_periodo
 
    #dadosquefaltam = getAnalistaDoProjetoECpfCoordenador(codigo)
    #print(f'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{dadosquefaltam}')
    #sheet['H47'] = formatar_cpf(dadosquefaltam["CPF_COORDENADOR"])
    meses_dict = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}   
    strintT = tamanho
    stringTamanho = f'A{tamanho}' # retorna lugar de brasilia
    hoje = date.today()
    data_formatada = f"{hoje.day} de {meses_dict[hoje.month]} de {hoje.year}"
    sheet[stringTamanho] = f'Brasilia, {data_formatada}'


    #Preenchendo a planilha
    # print("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEXEC")
    #print(dicionarioComDadosEntreAsDatas)
    rowInicial = 16
    for key,value in dicionarioComDadosEntreAsDatas.items():
        if key == "Encargos - ISS 5% ":
              sheet['C17'] = key
              sheet['E17'] = value
        elif key == "Encargos - ISS 2% ":
              sheet['C18'] = key
              sheet['E18'] = value
        else:
            sheet[f"H{rowInicial}"] = key
            sheet[f'I{rowInicial}'] = value
    
        rowInicial = rowInicial+1
             
    # print(dicionarioComDadosEntreAsDatasComMaterial)
    
    for key,value in dicionarioComDadosEntreAsDatasComMaterial.items():
        if key == "Receitas":
              sheet['C16'] = key
              sheet['E16'] = value
        elif key == 'Equipamento Material Permanente':
              sheet[f'I{tamanhoequipamentos}'] = value
        elif key == 'Material Permanente e Equipamento Nacional':
              sheet[f'I{tamanhoequipamentos+1}'] = value
        elif key == 'Material Permanente e Equipamento Importado':
              sheet[f'I{tamanhoequipamentos+2}'] = value
        elif key == 'Rendimentos de Aplicações Financeiras':
             sheet[f'A{tamanhoequipamentos+6}'] = 'RF Ref DI Plus Ágil '
             sheet[f'E{tamanhoequipamentos+6}'] = value

              

    workbook.save(planilha)
    workbook.close()

    return strintT
    
def execReceitaDespesa(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho,variavelResumo,variavelResumoTotalAteadata,valorPrevisto):
    
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Exec. Receita e Despesa")
    workbook.save(tabela)
    workbook.close()




    ####VALOR PREVISTO
    
    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',valorPrevisto)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',valorPrevisto)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',valorPrevisto)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',valorPrevisto)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',valorPrevisto)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',valorPrevisto)
    #rendimentodeapliaçõesfinanceiras
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',valorPrevisto)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',valorPrevisto)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',valorPrevisto)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',valorPrevisto) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',valorPrevisto) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',valorPrevisto)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',valorPrevisto)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',valorPrevisto)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',valorPrevisto)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',valorPrevisto)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',valorPrevisto)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Importado',valorPrevisto)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',valorPrevisto)

    # print(dict14)

    dicionarioValorPrevisto = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    mergeTotalPrevisto = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19}

    dicionarioValoresPrevisto = {}
    for key, value in dicionarioValorPrevisto.items() :
        if value != 0: 
            dicionarioValoresPrevisto[key]= value 




    ##########PREENCHE COM BETWEEN - REALIZADO EXECTUADO NO PERIODO
    #pessoal fisica
    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',variavelResumo)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',variavelResumo)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',variavelResumo)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',variavelResumo)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',variavelResumo)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',variavelResumo)
    #rendimentodeapliaçõesfinanceiras
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',variavelResumo)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',variavelResumo)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',variavelResumo)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',variavelResumo) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',variavelResumo) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',variavelResumo)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',variavelResumo)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',variavelResumo)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',variavelResumo)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',variavelResumo)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',variavelResumo)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Nacional',variavelResumo)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',variavelResumo)
    #Receitas
    dict20 = resumoReceitaDespesa(2,rubricas,'Receitas',variavelResumo)
    #Despesas financeiras
    dict21 = resumoReceitaDespesa(9,rubricas,'Despesas Financeiras',variavelResumo)



    dicionarioQueVaiSerPrintado = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    merged_dict = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19,**dict20,**dict21}




    dictPraCalcularTamanho = {}
    for key, value in dicionarioQueVaiSerPrintado.items() :
        if value != 0: 
            dictPraCalcularTamanho[key]= value 
    
    newDictValoresPrevisto = dict(dicionarioValoresPrevisto)
    newDictPraCalcularTamanho = dict(dictPraCalcularTamanho)
    
    newDictPraCalcularTamanho.update(dicionarioValoresPrevisto)
    newDictPraCalcularTamanho.update(dictPraCalcularTamanho)

    tamanho = len(newDictPraCalcularTamanho)
    #print(dictPraCalcularTamanho)
    
    stringTamanho = tamanho + 16 
    estiloExecReceitaDespesa(tabela,tamanho,stringTamanho)


    ##############SEM O BETWEEN REALIZADO ACUMULADO ATE O PERIODO

    rubricasFisica = [79,54,55]
    dict1 = resumoReceitaDespesa(25,rubricasFisica,'Outros Serviços de Terceiros - Pessoa Física',variavelResumoTotalAteadata)
    #pessoa Juridica
    rubricasJuridica = [57,26]
    dict2 = resumoReceitaDespesa(75,rubricasJuridica,'Outros Serviços de Terceiros - Pessoa Jurídica',variavelResumoTotalAteadata)
    #passagemLocomoção
    rubricasPassagem = [52,20]
    dict3 = resumoReceitaDespesa(7,rubricasPassagem,'Passagens e Despesas com Locomoção',variavelResumoTotalAteadata)
     #serv.terceiro celetista
    rubricasCeletista = [87,68,69,70]
    dict4 = resumoReceitaDespesa(156,rubricasCeletista,'Serv. Terceiro Celetistas',variavelResumoTotalAteadata)
    #Obricaçãotributaria 20% de OST
    rubricas =[]
    dict5 = resumoReceitaDespesa(86,rubricas,f'Obrigações Tributárias e Contributivas - 20% de OST',variavelResumoTotalAteadata)
    #Obricaçãotributaria
    dict6 = resumoReceitaDespesa(66,rubricas,f'Obrigações Tributárias e contributivas',variavelResumoTotalAteadata)
    #Rendimento de aplicação
    dict7 = resumoReceitaDespesa(3,rubricas,f'Rendimentos de Aplicações Financeiras',variavelResumoTotalAteadata)
    #diarias
    rubricasDiarias = [37,50,10,78,30,51,63,11]
    dict8 = resumoReceitaDespesa(53,rubricasDiarias,'Diárias',variavelResumoTotalAteadata)
    #auxilio
    rubricasAuxilio = [32]
    dict9 = resumoReceitaDespesa(21,rubricasAuxilio,'Auxílio Financeiro',variavelResumoTotalAteadata)
    #BolsaExtensão
    dict10 = resumoReceitaDespesa(96,rubricas,'Bolsa Extensão',variavelResumoTotalAteadata) 
    #estagiairio
    rubricasEstagiario = [56]  
    dict11 = resumoReceitaDespesa(80,rubricasEstagiario,'Estagiário',variavelResumoTotalAteadata) 
    #custo indireto
    rubricasCustos = [77,111,117]
    dict12 = resumoReceitaDespesa(49,rubricasCustos,'Custos Indiretos - FUB',variavelResumoTotalAteadata)
    #Relaçãodebens
    dict13 = resumoReceitaDespesa(112,rubricas,'Relação de Bens',variavelResumoTotalAteadata)
    #Material de Consumo
    rubricasMaterialConsumo = [16,15]
    dict14 = resumoReceitaDespesa(47,rubricasMaterialConsumo,'Material de Consumo',variavelResumoTotalAteadata)
    #equipamentoMaterial
    rubricasEquipamentoMaterial = [17,18]
    dict15 = resumoReceitaDespesa(39,rubricasEquipamentoMaterial,'Equipamento Material Permanente',variavelResumoTotalAteadata)
    #ISS 2%
    rubricas =[]
    dict16 = resumoReceitaDespesa(88,rubricas,'Encargos - ISS 2% ',variavelResumoTotalAteadata)
    #ISS 5%
    dict17 = resumoReceitaDespesa(67,rubricas,'Encargos - ISS 5% ',variavelResumoTotalAteadata)
    #equipamentoMaterialNacional
    dict18 = resumoReceitaDespesa(17,rubricas,'Material Permanente e Equipamento Nacional',variavelResumoTotalAteadata)
    #equipamentoMaterialInternacional
    dict19 = resumoReceitaDespesa(18,rubricas,'Material Permanente e Equipamento Importado',variavelResumoTotalAteadata)



    dicionarioAcumuladoAteOPeriodo = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14,**dict16,**dict17,**dict18,**dict19}


    mergeTotal = {**dict1, **dict2, **dict3, **dict4, **dict5, **dict6, **dict7, **dict8, **dict9, **dict10, **dict11, **dict12, **dict13, **dict14, **dict15,**dict16,**dict17,**dict18,**dict19}

    dicionarioValoresTotais = {}
    for key, value in dicionarioAcumuladoAteOPeriodo.items() :
        if value != 0: 
            dicionarioValoresTotais[key]= value 
    

    #preencher
    workbook = openpyxl.load_workbook(planilha)
    sheet = workbook['Exec. Receita e Despesa']

    input_date = []
    output_date_str = []
    input_date2  = []
    output_date_str2 = []
    if check_format(data1):
        input_date = datetime.strptime(data1, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str = input_date.strftime("%d/%m/%Y")
    else :
         return None
    if check_format(data2):
        input_date2 = datetime.strptime(data2, "%Y-%m-%d")
    # Format the datetime object to a string in dd/mm/yyyy format
        output_date_str2 = input_date2.strftime("%d/%m/%Y")
    else :
         return None
    
    
    
    string_periodo = f"Período que abrange esta prestação: {output_date_str} a {output_date_str2}"
    sheet['A7'] = string_periodo
    consulta_coordenador = consultaID(codigo)
    # print(consulta_coordenador)
    # print(type(consulta_coordenador))
    stringCoordenador= f'F{stringTamanho+11}' # retorna lugar do coordanor
    stringCoordanadorCargo = f'F{stringTamanho+12}'
    sheet[stringCoordanadorCargo] = f"Coordenador(a)"
    stringTamanhoCPF = f'F{stringTamanho+13}' # retorna lugar do coordanor
    sheet[stringCoordenador] = consulta_coordenador['NomePessoaResponsavel']
    sheet[stringTamanhoCPF] = formatar_cpf(consulta_coordenador['CPFCoordenador'])
    string_titulo = f"Título do Projeto: {consulta_coordenador['NomeConvenio']}"
    string_executora = f"Executora: {consulta_coordenador['NomePessoaFinanciador']}"
    string_participe = f"Partícipe: Fundação de Empreendimentos Científicos e Tecnológicos - FINATEC"
   # Convert 'DataAssinatura' to "dd/mm/YYYY" format
    datetime_obj1 = consulta_coordenador['DataAssinatura']
    formatted_date1 = datetime_obj1.strftime("%d/%m/%Y")

    # Convert 'DataVigencia' to "dd/mm/YYYY" format
    datetime_obj2 = consulta_coordenador['DataVigencia']
    formatted_date2 = datetime_obj2.strftime("%d/%m/%Y")

# Create the string representing the period of execution
    string_periodo = f"Período de Execução Físico-Financeiro: {formatted_date1} a {formatted_date2}"
    sheet['A3'] = string_titulo
    sheet['A4'] = string_executora
    sheet['A5'] = string_participe
    sheet['A6'] = string_periodo
 
    #dadosquefaltam = getAnalistaDoProjetoECpfCoordenador(codigo)
    #print(f'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa{dadosquefaltam}')
    #sheet['H47'] = formatar_cpf(dadosquefaltam["CPF_COORDENADOR"])
    meses_dict = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}   

    stringTamanhoBrasilia = f'A{stringTamanho+10}' # retorna lugar de brasilia
    hoje = date.today()
    data_formatada = f"{hoje.day} de {meses_dict[hoje.month]} de {hoje.year}"
    sheet[stringTamanhoBrasilia] = f'Brasilia, {data_formatada}'

    #despesas correntes printr
    imprimirResumoLinha = 16


    # print("dictpracalcularTamanho")
    # print(dictPraCalcularTamanho)
    # print("dictprevisto")
    # print(dicionarioValoresPrevisto)



    #REALIZADO
    for key,value in newDictPraCalcularTamanho.items():
          if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringChave = f"A{imprimirResumoLinha}"
            imprimirResumoLinha = 1 + imprimirResumoLinha
            sheet[posicaoStringChave]=key


    imprimirResumoLinha = 16

    for key,value in dictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"C{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    posicaoEquipamentoMaterialPermanente = 16 + tamanho + 3
    for key,value in merged_dict.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"C{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value

    #REALIZADO ACUMULADO ATE O PERIODO
    imprimirResumoLinha = 16
    for key,value in dicionarioValoresTotais.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"G{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    for key,value in mergeTotal.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"G{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value

 
             
    #valor previsto            
    imprimirResumoLinha = 16
    newDictPraCalcularTamanho.update(dicionarioValoresPrevisto)
    for key,value in newDictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"B{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha

    imprimirResumoLinha = 16
    for key,value in newDictPraCalcularTamanho.items():
        if key != "Despesas Financeiras" and key != "Equipamento Material Permanente":
            posicaoStringValor = f"F{imprimirResumoLinha}"
            sheet[posicaoStringValor]=value
            imprimirResumoLinha = 1 + imprimirResumoLinha


    for key,value in mergeTotalPrevisto.items():
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"B{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value
        if key == "Equipamento Material Permanente":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Nacional":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente+1}"
            sheet[posicaoStringValor]=value
        if key == "Material Permanente e Equipamento Importado":
            posicaoStringValor = f"F{posicaoEquipamentoMaterialPermanente+2}"
            sheet[posicaoStringValor]=value



    workbook.save(planilha)
    workbook.close()
    return tamanho,dictPraCalcularTamanho,merged_dict

def pessoaFisica(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
   
    tabela = pegar_caminho(planilha)
    nomeTabela ="Outros Serviços Terceiros - PF"
    tituloStyle = "pessoaFisica"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Outros Serviços Terceiros - PF")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    

    tamanho = []
   
    rubricas = [79,54,55]
    if 25 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[25].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 25 not in categorized_data:
                    categorized_data[25] = categorized_data[num]
                else:
                    categorized_data[25].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 25 not in categorized_data or not categorized_data[25]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[25])
    print(maior)
    print(len(categorized_data[25]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet555 = workb['Outros Serviços Terceiros - PF']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet555.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 25 not in categorized_data or not categorized_data[25]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[25],li)
        
    
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet555.cell(row=rowkek, column=coluna, value=cell_data)
            # print(cell_data)
            # print(f'row :')
            # print(rowkek)
            # print(f'coluna :')
            # print(coluna)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    #print(valores_preenchimento)
    workb.save(tabela)
    workb.close()

# ##########################################Pessoa Juridica#########################################
def pessoaJuridica(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
   
   

    tabela = pegar_caminho(planilha)
    nomeTabela ="Pessoa Jurídica"
    tituloStyle = "pessoaJuridica"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Pessoa Jurídica")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica


    tamanho = []
   
    rubricas = [57,26]
    if 75 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[75].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 75 not in categorized_data:
                    categorized_data[75] = categorized_data[num]
                else:
                    categorized_data[75].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 75 not in categorized_data or not categorized_data[75]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[75])
    print(maior)
    print(len(categorized_data[75]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet5 = workb['Pessoa Jurídica']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet5.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 75 not in categorized_data or not categorized_data[75]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[75],li)
        
        # n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet5.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(tabela)
    workb.close()

# ##########################################ISS#########################################CANCELADO
def iss(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    nomeTabela ="ISS"
    tituloStyle = "isss"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="ISS")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica

    rubricas = [88]
    if 67 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[67].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 67 not in categorized_data:
                    categorized_data[67] = categorized_data[num]
                else:
                    categorized_data[67].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 67 not in categorized_data or not categorized_data[67]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[67])
    print(maior)
    print(len(categorized_data[67]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet3 = workb["ISS"]

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet3.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 67 not in categorized_data or not categorized_data[67]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[67],li)
        
        # n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet3.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    

    workb.save(tabela)
    workb.close()

# #########################################Passagem Locomoção#########################################
def passagemLocomocao(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    nomeTabela ="Passagens e Desp. Locomoção"
    tituloStyle = "passagenDespLocomo"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Passagens e Desp. Locomoção")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica

    rubricas = [52,20]
    if 7 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[7].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 7 not in categorized_data:
                    categorized_data[7] = categorized_data[num]
                else:
                    categorized_data[7].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 7 not in categorized_data or not categorized_data[7]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[7])
    print(maior)
    print(len(categorized_data[7]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet3 = workb["Passagens e Desp. Locomoção"]

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet3.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 7 not in categorized_data or not categorized_data[7]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[7],li)
        
        # n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet3.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    

    workb.save(tabela)
    workb.close()

# ##########################################Serv.Terceiro CLTa#########################################CANCELADO
def terclt(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    nomeTabela ="Serv. Terceiro CLT"
    tituloStyle = "outrosServiçosTerceiros"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Serv. Terceiro CLT")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    
    
    rubricas = [87,68,69,70]
    if 156 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[156].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 156 not in categorized_data:
                    categorized_data[156] = categorized_data[num]
                else:
                    categorized_data[156].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 156 not in categorized_data or not categorized_data[156]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[156])
    print(maior)
    print(len(categorized_data[156]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet3 = workb["Serv. Terceiro CLT"]

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet3.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 156 not in categorized_data or not categorized_data[156]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[156],li)
        
        # n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet3.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    

    workb.save(tabela)
    workb.close()
# ##########################################Obrigaçoes tributárias #########################################
def obricacaoTributaria(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    nomeTabela ="Obrigações Trib. - Encargos 20%"
    tituloStyle = "obrigacoesTribu"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Obrigações Trib. - Encargos 20%")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    rubricas = [86]
    if 66 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[66].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 66 not in categorized_data:
                    categorized_data[66] = categorized_data[num]
                else:
                    categorized_data[66].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 66 not in categorized_data or not categorized_data[66]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[66])
    print(maior)
    print(len(categorized_data[66]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet4 = workb['Obrigações Trib. - Encargos 20%']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet4.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 66 not in categorized_data or not categorized_data[66]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)

                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[66],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet4.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(planilha)
    workb.close()

# ##########################################Conciliação Bancária #########################################
# Função para converter o mês para o formato desejado
def formatar_data(row):
    dia = row.day
    mes = row.month
    ano = row.year

    # Mapear o número do mês para o nome abreviado
    meses = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}

    # Obter o nome abreviado do mês
    mes_abreviado = meses.get(mes, mes)

    # Criar a string formatada
    data_formatada = f'{dia}-{mes_abreviado}-{ano}'
    
    return data_formatada
def conciliacaoBancaria(codigo,data1,data2,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Conciliação Bancária")
    workbook.save(tabela)
    workbook.close()
    tamanho = []
    # categorized_data= separarporrubrica(codigo,data1,data2)

    dataframeSemEstorno,dataframeComEstorno = consultaConciliacaoBancaria(codigo,data1,data2)
   
    #####pergar os dados do db e separar por mes e ano###################3
    
    grupos_por_ano_mes = defaultdict(list)
    if  dataframeSemEstorno.empty and dataframeComEstorno.empty:
                print("Data not available or empty.")
                maior = 5
                maior2= 5
                tabela = pegar_caminho(planilha)
                estilo_conciliacoes_bancaria(tabela,maior,maior2,stringTamanho)
                return None  # or handle the case accordingly
    else:
        
        tamanho = len(dataframeSemEstorno)
        tamanho2 = len(dataframeComEstorno)

        tabela = pegar_caminho(planilha)
        estilo_conciliacoes_bancaria(tabela,tamanho,tamanho2,stringTamanho)
       

        workb = openpyxl.load_workbook(tabela)
        worksheet333 = workb["Conciliação Bancária"]
        i = 16
        j=0
        estorno_valor = 0
        
        dataframeSemEstorno['data_formatada'] = dataframeSemEstorno['DataPagamento'].apply(formatar_data)
        dataframeComEstorno['data_formatada'] = dataframeComEstorno['DataPagamento'].apply(formatar_data)
        dataframeSemEstorno['DataPagamento'] = dataframeSemEstorno['data_formatada']
        dataframeComEstorno['DataPagamento'] = dataframeComEstorno['data_formatada']

        dataframeSemEstorno = dataframeSemEstorno.drop('data_formatada', axis=1)
        dataframeComEstorno = dataframeComEstorno.drop('data_formatada', axis=1)

        #for row in worksheet333.iter_rows(min_row=16, max_row=tamanho, min_col=1, max_col=4):

       # Write data starting from the first row
        for row_num, row_data in enumerate(dataframeSemEstorno.itertuples(index=False), start=16):#inicio linha
            for col_num, value in enumerate(row_data, start=1):#inicio coluna
                worksheet333.cell(row=row_num, column=col_num, value=value)
                print(row_num)
                print(col_num)
       
        linha2 = 16+4+tamanho


        for row_num, row_data in enumerate(dataframeComEstorno.itertuples(index=False), start=linha2):#inicio linha
            for col_num, value in enumerate(row_data, start=1):#inicio coluna
                worksheet333.cell(row=row_num, column=col_num, value=value)
                print("comestorno")
                print(row_num)
                print(col_num)
        

        workb.save(tabela)
        workb.close

# ##########################################Rendimento de Aplicação#########################################
def rendimentodeaplicacao(codigo,data1,data2,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Rendimento de Aplicação")
    workbook.save(tabela)
    workbook.close()
    tamanho = []
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    #####pergar os dados do db e separar por mes e ano###################3
    
    grupos_por_ano_mes = defaultdict(list)
    if 3 not in categorized_data or not categorized_data[3]:
                print("Data not available or empty.")
                maior = 1
                maior2= 2
                tabela = pegar_caminho(planilha)
                estilo_rendimento_de_aplicacao(tabela,maior,stringTamanho)
                return None  # or handle the case accordingly
    else:
        
        for item in categorized_data[3]:
            data_criacao_str = item['DataCriacao']
            
            # Converter a string de data para um objeto datetime
            data_criacao = datetime.strptime(data_criacao_str, '%d/%m/%Y')
            # Extrair o componente do ano e do mês
            ano = data_criacao.year
            mes = data_criacao.month
            dia = data_criacao.day
            # Adicionar o item ao grupo correspondente ao ano e mês
                
            grupos_por_ano_mes[(ano, mes,dia)].append(item)

            # Calcular a soma de VALOR_LANCADO e imprimir os resultados
            
        estorno = defaultdict(list)
        
        tamanho = len(grupos_por_ano_mes)     
        tabela = pegar_caminho(planilha)
        #print(tabela)
        estilo_rendimento_de_aplicacao(tabela,tamanho,stringTamanho)
       

        workb = openpyxl.load_workbook(tabela)
        worksheet344 = workb["Rendimento de Aplicação"]
        i = 14
       
        for (ano, mes,dia), items in sorted(grupos_por_ano_mes.items()):  
            soma_valor_lancado = 0
            for item in items:
                soma_valor_lancado += item['ValorLancamento']


            anoss = {1:'jan',
                2:'fev',
                3:'mar',
                4:'abr',
                5:'mai',
                6:'jun',
                7:'jul',
                8:'ago',
                9:'sep',
                10:'out',
                11:'nov',
                12: 'dec'
                    
            }
            for a,b in anoss.items():
                if mes == a :
                    mes = b
            cell_data = f'{mes}-{ano}'
            # print(cell_data)
            # print(valor_lancado)
            
            worksheet344.cell(row=i, column=1, value=cell_data)
            worksheet344.cell(row=i,column=8,value=soma_valor_lancado)
           
            i = i + 1
           
         
      
        workb.save(tabela)
        workb.close
   ##############################

def diaria(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Diárias"
    tituloStyle = "diarias"
    sheet2 = workbook.create_sheet(title="Diárias")
    workbook.save(tabela)
    workbook.close()
    
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica

    tamanho = []
    #diarias
    #53,37,50,10,
    rubricas = [37,50,10,78,30,51,63,11]
    if 53 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[53].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 53 not in categorized_data:
                    categorized_data[53] = categorized_data[num]
                else:
                    categorized_data[53].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 53 not in categorized_data or not categorized_data[53]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[53])
    print(maior)
    print(len(categorized_data[53]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet5 = workb['Diárias']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet5.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 53 not in categorized_data or not categorized_data[53]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[53],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet5.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(tabela)
    workb.close()

def auxilio(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    nomeTabela ="Auxílio Financeiro"
    tituloStyle = "auxilio"
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Auxílio Financeiro")
    workbook.save(tabela)
    workbook.close()
    
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    rubricas = [32]
    if 31 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[31].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 31 not in categorized_data:
                    categorized_data[31] = categorized_data[num]
                else:
                    categorized_data[31].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 31 not in categorized_data or not categorized_data[31]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[31])
    print(maior)
    print(len(categorized_data[31]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet4 = workb['Auxílio Financeiro']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet4.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 31 not in categorized_data or not categorized_data[31]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)

                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[31],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet4.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(planilha)
    workb.close()

def bolsaExtensao(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Bolsa Extensão"
    tituloStyle = "bolsaExtensao"
    sheet2 = workbook.create_sheet(title="Bolsa Extensão")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    rubricas = []
    if 96 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[96].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 96 not in categorized_data:
                    categorized_data[96] = categorized_data[num]
                else:
                    categorized_data[96].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 96 not in categorized_data or not categorized_data[96]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[96])
    print(maior)
    print(len(categorized_data[96]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet3 = workb["Bolsa Extensão"]

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet3.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 4 not in categorized_data or not categorized_data[4]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[4],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet3.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    

    workb.save(tabela)
    workb.close()

def estagiario(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Estagiário"
    tituloStyle = "estagiario"
    sheet2 = workbook.create_sheet(title="Estagiário")
    workbook.save(tabela)
    workbook.close()
   
    categorized_data = dadoRubrica
    rubricas = [56]
    if 80 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[80].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 80 not in categorized_data:
                    categorized_data[80] = categorized_data[num]
                else:
                    categorized_data[80].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 80 not in categorized_data or not categorized_data[80]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[80])
    print(maior)
    print(len(categorized_data[80]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet4 = workb['Estagiário']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet4.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 80 not in categorized_data or not categorized_data[80]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)

                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[80],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet4.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(planilha)
    workb.close()

def custoIndireto(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Custos Indiretos - FUB"
    tituloStyle = "custosIndiretos"
    sheet2 = workbook.create_sheet(title="Custos Indiretos - FUB")
    workbook.save(tabela)
    workbook.close()
    categorized_data = dadoRubrica
    rubricas = [77,111,117]
    if 49 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[49].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 49 not in categorized_data:
                    categorized_data[49] = categorized_data[num]
                else:
                    categorized_data[49].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 49 not in categorized_data or not categorized_data[49]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[49])
    print(maior)
    print(len(categorized_data[49]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet4 = workb['Custos Indiretos - FUB']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet4.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 49 not in categorized_data or not categorized_data[49]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)

                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[49],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet4.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(planilha)
    workb.close()

def relacaodeBens(codigo,data1,data2,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Relação de Bens"
    tituloStyle = "relacaoBEns"
    sheet2 = workbook.create_sheet(title="Relação de Bens")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica
    rubricas = []
    if 112 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[112].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 112 not in categorized_data:
                    categorized_data[112] = categorized_data[num]
                else:
                    categorized_data[112].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 15
        tabela = pegar_caminho(planilha)
        estiloRelacaoBens(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    if 112 not in categorized_data or not categorized_data[112]:
                maior = 15
                tabela = pegar_caminho(planilha)
                estiloRelacaoBens(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[112])
    print(maior)
    print(len(categorized_data[112]))
    tabela = pegar_caminho(planilha)

    estiloRelacaoBens(tabela,maior,tituloStyle,nomeTabela,stringTamanho)

    #### falta preencher

def materialDeConsumo(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Material de Consumo"
    tituloStyle = "materialDeConsumo"
    sheet2 = workbook.create_sheet(title="Material de Consumo")
    workbook.save(tabela)
    workbook.close()

    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica

    rubricas = [16,15]
    if 47 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[47].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 47 not in categorized_data:
                    categorized_data[47] = categorized_data[num]
                else:
                    categorized_data[47].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
     
    if 47 not in categorized_data or not categorized_data[47]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[47])
    print(maior)
    print(len(categorized_data[47]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet3 = workb["Material de Consumo"]

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet3.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 47 not in categorized_data or not categorized_data[47]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[47],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet3.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1
    

    workb.save(tabela)
    workb.close()

def equipamentoMaterialPermanente(codigo,data1,data2,keys,planilha,dadoRubrica,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Equipamento Material Permanente"
    tituloStyle = "equipamentoMaterialPermanente"
    sheet2 = workbook.create_sheet(title="Equipamento Material Permanente")
    workbook.save(tabela)
    workbook.close()
    # categorized_data= separarporrubrica(codigo,data1,data2)
    categorized_data = dadoRubrica

    tamanho = []
   
    # if 39 in categorized_data and 18 in categorized_data:
    #     categorized_data[39].extend(categorized_data[18])
    # elif 39 not in categorized_data and 18 in categorized_data:
    #    categorized_data[39] = categorized_data[18]
    rubricas = [17,18]
    if 39 in categorized_data:
        for num in rubricas:
            if num in categorized_data:
                    categorized_data[39].extend(categorized_data[num])
    elif any(num in categorized_data for num in rubricas):
        for num in rubricas:
            if num in categorized_data:
                if 39 not in categorized_data:
                    categorized_data[39] = categorized_data[num]
                else:
                    categorized_data[39].extend(categorized_data[num])
    else:
        print("Data not available or empty.")
        maior = 1
        tabela = pegar_caminho(planilha)
        estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
        return None  # or handle the case accordingly
    
    if 39 not in categorized_data or not categorized_data[39]:
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                print("Data not available or empty.")
                return None  # or handle the case accordingly
    maior = len(categorized_data[39])
    print(maior)
    print(len(categorized_data[39]))
    tabela = pegar_caminho(planilha)

    estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)


    coluna = 2
    # caminho = pegar_caminho(planilha)

    workb = openpyxl.load_workbook(tabela)
    worksheet5 = workb['Equipamento Material Permanente']

    for i in range(1,maior+1):
        valor_coluna = 9 + i
        worksheet5.cell(row=valor_coluna, column=1, value=i)  # column index starts from 1


    for i in keys:
        li = [i]
        if 39 not in categorized_data or not categorized_data[39]:
                print("Data not available or empty.")
                maior = 1
                tabela = pegar_caminho(planilha)
                estiloGeral(tabela,maior,tituloStyle,nomeTabela,stringTamanho)
                return None  # or handle the case accordingly
        valores_preenchimento = retornavalores(categorized_data[39],li)
        
        n = len(valores_preenchimento)  
        for rowkek, cell_data in enumerate(valores_preenchimento, start=10):
            worksheet5.cell(row=rowkek, column=coluna, value=cell_data)  
        # if coluna == 5 or coluna == 7 :
        #         coluna = coluna + 1  
        coluna = coluna + 1

    workb.save(tabela)
    workb.close()

def demonstrativo(codigo,data1,data2,planilha,tamanco):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Demonstrativo de Receita")
    workbook.save(tabela)
    workbook.close()
    tamanho = 20
    estilo_demonstrativoDeReceita(tabela,tamanho,tamanco)

def preencher_fub_teste(codigo,data1,data2,keys,tabela):
    dadoRubrica= separarporrubrica(codigo,data1,data2)
    variavelResumoComPeriodo = queryReceitaXDespesa(codigo,data1,data2)
    variavelResumoTotalAteadata = queryReceitaXDespesaTotal(codigo,data2)
    variavelResumoPrevisto = queryValorPrevisto(codigo)
    #print(variavelResumoPrevisto)
    #print(dadoRubrica)
    #tamanhoPosicaoBrasilia = criaReceitaXDespesa(tabela,codigo,data1,data2,300)
    tamanhoResumoExec,dicionarioComDadosEntreAsDatas,dicionarioComMaterial= execReceitaDespesa(codigo,data1,data2,keys,tabela,dadoRubrica,15,variavelResumoComPeriodo,variavelResumoTotalAteadata,variavelResumoPrevisto)
    #chamar receitaxdepesa para finalizala
    tamanhoPosicaoBrasilia = criaReceitaXDespesa(tabela,codigo,data1,data2,tamanhoResumoExec,dicionarioComDadosEntreAsDatas,dicionarioComMaterial)
    #preencherCapa(codigo,tabela)
    pessoaFisica(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    pessoaJuridica(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    #iss(codigo,data1,data2,keys,tabela,dadoRubrica)
    passagemLocomocao(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    terclt(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    obricacaoTributaria(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    conciliacaoBancaria(codigo,data1,data2,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    rendimentodeaplicacao(codigo,data1,data2,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    diaria(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    auxilio(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    bolsaExtensao(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    estagiario(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    custoIndireto(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    materialDeConsumo(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    equipamentoMaterialPermanente(codigo,data1,data2,keys,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
    demonstrativo(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    relacaodeBens(codigo,data1,data2,tabela,dadoRubrica,tamanhoPosicaoBrasilia)
   
    
    

# keys = ['NOME_FAVORECIDO','CNPJ_FAVORECIDO','TIPO_LANCAMENTO','HIS_LANCAMENTO','DATA_EMISSAO','DATA_PAGAMENTO', 'VALOR_PAGO']
# tabela = pegar_caminho("Modelo_Fub.xlsx")
# preencher_fub_teste(6411,'2020-01-01','2024-01-31',keys,tabela)

# pessoa_fisica(6858,'2022-09-09','2022-12-09',keys)
