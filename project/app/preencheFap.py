import pyodbc
from datetime import datetime,date
import openpyxl
from openpyxl.styles import Font
import os
from .estilo_fub import *
from collections import defaultdict
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import numpy as np  

def convert_datetime_to_string(value):
    if isinstance(value, datetime):
        return value.strftime('%d/%m/%Y')
    return value

def convert_datetime_to_stringdt(dt):
    # Check if the value is a pandas Timestamp
    if isinstance(dt, pd.Timestamp):
        # Convert the Timestamp to a string using strftime
        return dt.strftime('%d/%m/%Y')  # You can customize the format as needed
    else:
        # If it's not a Timestamp, return the original value
        return dt

def formatar_data(row):
    """ Formata a data com o mes abreviado transformando 01 em jan por exemplo
    """
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

def formatarDataSemDia(row):
    """ Formata a data com o mes abreviado transformando 01 em jan por exemplo
    """
    dia = row.day
    mes = row.month
    ano = row.year

    # Mapear o número do mês para o nome abreviado
    meses = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}

    # Obter o nome abreviado do mês
    mes_abreviado = meses.get(mes, mes)

    # Criar a string formatada
    data_formatada = f'{mes_abreviado}-{ano}'
    
    return data_formatada

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


    # return records

#todas as consultas em sql
def consultaAnexoUm(IDPROJETO):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO,)]
    queryNomeConvenioProcessoSubProcesso = f"SELECT [LisConvenio].NomeConvenio ,[LisConvenio].Processo,SubProcesso,ValorAprovado FROM [Conveniar].[dbo].[LisConvenio] WHERE CodConvenio = ? "
    dfConvenioProcessoSubProcessos = pd.read_sql(queryNomeConvenioProcessoSubProcesso, engine, params=parametros)

    
    return dfConvenioProcessoSubProcessos

def consultaAnexoDois(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO,DATA1,DATA2)]
    queryAnexoDois = f"""SELECT [NomeRubrica]
		,NumChequeDeposito
		,CONVERT(varchar(10), DataPagamento, 103) AS FormattedDate
		,NumDocPago
		,[NomeFavorecido]
		,HisLancamento
		,ValorPago
        FROM [Conveniar].[dbo].[LisLancamentoConvenio]
     WHERE [LisLancamentoConvenio].CodConvenio = ? AND [LisLancamentoConvenio].CodStatus = 27
     AND [LisLancamentoConvenio].DataPagamento BETWEEN ? AND ? and [LisLancamentoConvenio].CodRubrica not in (2,3,9,67,88) order by DataPagamento"""
    dfConvenioAnexoDois = pd.read_sql(queryAnexoDois, engine, params=parametros)

    
    
    return dfConvenioAnexoDois

def consultaAnexoTres(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passss.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    idprojetoComZero = f"0{IDPROJETO}"
    parametros = [(IDPROJETO,idprojetoComZero, DATA1, DATA2)]
    
    queryConsultaComRubrica = f"""SELECT 
    [Data de Aquisição][dataAqui],
	[Nº Nota][nota],
    [Descrição][descri],
	[Valor de Aquisição][valorAqui],
    [Valor de Aquisição][valorAqui2],
    [Patrimônio][patri],
    [Localização][localiza],
    [Responsável][responsavel]
    FROM [SBO_FINATEC].[dbo].[VW_BENS_ADQUIRIDOS] 
    WHERE ([Cod Projeto] = ? or [Cod Projeto] = ? ) 
    AND [Status] = 'Imobilizado' 
    AND [Data de Aquisição] BETWEEN ? AND ? 
    Order by [Data de Aquisição]"""
    dfConsultaBens = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)

    
    return dfConsultaBens  
    return 0

def consultaRendimentosIRRFConciliacao(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    
    consultaComPeriodo =f"SELECT SUM(ValorPago) AS TotalPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 0 AND DataPagamento BETWEEN  ? AND ?"

    consultaRendimentoAplicacao = f"""
    SELECT 
        NomeTipoLancamento,
        SUM(CASE WHEN NomeTipoLancamento = 'IRRF Pessoa Jurídica' THEN ValorPago ELSE 0 END) AS IRRF,
        SUM(CASE WHEN NomeTipoLancamento = 'Aplicação Financeira' THEN ValorPago ELSE 0 END) AS Aplicação
    FROM 
        [Conveniar].[dbo].[LisLancamentoConvenio] 
    WHERE 
        CodConvenio = ? 
        AND CodStatus = 27 
        AND CodRubrica = 3 
        AND DataPagamento BETWEEN ? AND ?
    GROUP BY 
        NomeTipoLancamento;
    """


    Soma = pd.read_sql(consultaRendimentoAplicacao, engine, params=parametros)
   
    return Soma

def consultaDevolucaoRecursosConciliacao(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    
    consultaComPeriodo = f"SELECT SUM(ValorPago) AS TotalPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 0 AND DataPagamento BETWEEN  ? AND ?"


    Soma = pd.read_sql(consultaComPeriodo, engine, params=parametros)
   
    return Soma

def consultaConciliacao(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO,)]
    queryNomeConvenioProcessoSubProcesso = f"SELECT [LisConvenio].NomeConvenio ,[LisConvenio].Processo,SubProcesso,ValorAprovado FROM [Conveniar].[dbo].[LisConvenio] WHERE CodConvenio = ? "
    dfConvenioProcessoSubProcessos = pd.read_sql(queryNomeConvenioProcessoSubProcesso, engine, params=parametros)

    return dfConvenioProcessoSubProcessos


#preencher 

def anexoUm(tabela,codigo,data1,data2):
    consultaAnexoUm(codigo)
    return 0
def anexoDois(tabela,codigo,data1,data2):
    consultaAnexoDois(codigo,data1,data2)
def anexoTres(tabela,codigo,data1,data2):
    print(consultaAnexoTres(codigo,data1,data2))
    return 0
def Conciliacao(tabela,codigo,data1,data2):

       #rendimentosdeapliacação
    dfSoma = consultaRendimentosIRRFConciliacao(codigo,data1,data2)
    dfcComPeriodo = consultaDevolucaoRecursosConciliacao(codigo,data1,data2)
    Soma = dfSoma["Aplicação"] + dfSoma["IRRF"]
    
    all_null = Soma.isnull().all()
    if all_null != True :
            if len(Soma) == 1:
                result = Soma.iloc[0]
                print(f'resultado{result}')
            else:
                result = Soma.iloc[0] - Soma.iloc[1]
                print(result)
                # stringRendimento = f'Rendimento de Aplicação'
                # stringRendimentoValor = f'E{tamanhoequipamentos + 6}'
                # sheet[stringRendimentoValor] = result
                # sheet[f'A{tamanhoequipamentos + 6}'] = stringRendimento

    
    return 0

def preencheFap(codigo,data1,data2,tabela):
    '''Preencher fub legado
        dadoRubrica = transforma em dicionario a consulta feita por sql e separa por rubricas.
        variaveisResumo= preenchimento da tabela Exec, são definidas em tres tipos:
            -variavelResumoComPeriodo = consulta no sql com datas1 e data2
            -variavelResumoComPeriodo = consulta no sql ate o periodo da data2
            -variavelResumoComPeriodo = consulta no sql com valor total previsto
        execReceitaDespesa = preenche a sheet Exec.Receita e Despesa da planilha,
        e retorna 3 argumentos:
            -tamanho = o tamanho necessario para fazer o estilo, ele varia de acordo
            com a quantidade de rubricas
            -dictPraCalcularTamanho = o dicionario utilizado para calcular o tamanho, ele
            contem a rubricas utilizadas no projeto para poder preencher a sheet ReceitaxDespesa
            -merged_dict = dicionario que contem todas as rubricas inclusive equipamento e 
            material permanente, utilizado para colocar esses valores na ReceitaxDespesa
        

        Argumentos:
            codigo = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
            KEYS = Lista responsavel por filtrar no dicionario quais dados irão ser preenchidos nas
            planilhas, exemplo cpf,data etc
            tabela = tabela a ser preenchida  extensão xlsx


    # '''

    anexoUm(tabela,codigo,data1,data2)
    anexoTres(tabela,codigo,data1,data2)
    anexoDois(tabela,codigo,data1,data2)
    Conciliacao(tabela,codigo,data1,data2)

    # tamanho,dataframe = ExeReceitaDespesa(tabela,codigo,data1,data2,15)
    
    # tamanhoPosicaoBrasilia,dfReceitas,dfDemonstrativoReceitas = Receita(tabela,codigo,data1,data2,tamanho,dataframe)
    # demonstrativo(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia,dfDemonstrativoReceitas,dfReceitas)
    # # rubricaGeral(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    # #conciliacaoBancaria(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    # # rowRendimento= rendimentoDeAplicacao(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    # # relacaodeBens(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    
