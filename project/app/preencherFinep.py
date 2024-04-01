import pyodbc
from datetime import datetime,date
import openpyxl
from openpyxl.styles import Font
import os
from collections import defaultdict
from .estiloFINEP import *
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




#consulta

#ok
def consultaRelatorioExecFinanceiraA1(IDPROJETO, DATA1, DATA2):
    ''' Consulta que busca os valores executado no periodo, e os valores que foram executados até no periodo
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
    parametros2 = [(IDPROJETO, DATA2)]
    parametros3 = IDPROJETO
    consultaComPeriodo =f"SELECT NomeRubrica, SUM(ValorPago) AS VALOR_TOTAL_PERIODO FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica"
    consultaAteAData =  f"SELECT NomeRubrica, SUM(ValorPago) AS VALOR_TOTAL_DATA FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento <= ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica "
    #consultaPrevisto = f"SELECT NomeRubrica, SUM(VALOR*Quantidade) AS VALOR_TOTAL_PREVISTO FROM [Conveniar].[dbo].[LisConvenioItemAprovado] WHERE CodConvenio = ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica"
    
    dfComPeriodo= pd.read_sql(consultaComPeriodo, engine, params=parametros)
    dfAteAData = pd.read_sql(consultaAteAData, engine, params=parametros2)
    #dfPrevisto = pd.read_sql(consultaPrevisto, engine, params=(IDPROJETO,))

    return dfComPeriodo,dfAteAData
#ok
def consultaDemonstrativoReceitaEDespesaA2(IDPROJETO,DATA1,DATA2):
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
    parametros2 = [(IDPROJETO, DATA2)]
    parametros3 = IDPROJETO
    consultaComPeriodo =f"SELECT NomeRubrica, SUM(ValorPago) AS VALOR_TOTAL_PERIODO FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica"
    consultaAteAData =  f"SELECT NomeRubrica, SUM(ValorPago) AS VALOR_TOTAL_DATA FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento <= ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica "
    consultaPrevisto = f"SELECT NomeRubrica, SUM(VALOR*Quantidade) AS VALOR_TOTAL_PREVISTO FROM [Conveniar].[dbo].[LisConvenioItemAprovado] WHERE CodConvenio = ? GROUP BY NomeRubrica, CodRubrica Order by CodRubrica"
    consultaRubricaRecursoRecebidos =  f"SELECT sum(ValorPago) AS VALOR_TOTAL_DATA FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento <= ? "
    consultaRendimentoAplicacao = f"""
    SELECT 
        NomeTipoLancamento,
        --SUM(CASE WHEN NomeTipoLancamento = 'IRRF Pessoa Jurídica' THEN ValorPago ELSE 0 END) AS IRRF,
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
    dfComPeriodo= pd.read_sql(consultaComPeriodo, engine, params=parametros)
    dfAteAData = pd.read_sql(consultaAteAData, engine, params=parametros2)
    dfPrevisto = pd.read_sql(consultaPrevisto, engine, params=(IDPROJETO,))
    dfRubricaRecursoRecebidos = pd.read_sql(consultaRubricaRecursoRecebidos, engine, params=parametros2)



    return dfComPeriodo,dfAteAData,dfPrevisto,dfRubricaRecursoRecebidos,Soma
#ok
def consultaPagamentoPessoal(IDPROJETO,DATA1,DATA2):
    
    ''' Consulta dinamica do SQL relacionado a rubrica correspondente,cada pagina tem sua própria consulta correspondente a rubrica
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
            codigoRubrica = código da rubrica 
    '''
    codigoRubrica = 87
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    parametrosComRubricaEstorno  = [(IDPROJETO, DATA1, DATA2,IDPROJETO,DATA1, DATA2)]
    parametrosPJ=[(IDPROJETO, DATA1, DATA2)]
    parametrosPJestorno=[(IDPROJETO, DATA1, DATA2,IDPROJETO, DATA1, DATA2)]
    queryConsultaComRubrica = f"""SELECT NomeFavorecido,
     CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') 
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     NomeTipoLancamento,
     [LisConvenioItemAprovado].[DescConvenioItemAprovado],
     NumDocPago,
     DataEmissao,
     NumChequeDeposito,
     DataPagamento,
     ValorPago 
     FROM [Conveniar].[dbo].[LisLancamentoConvenio] 
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado] 
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado] 
     WHERE [LisLancamentoConvenio].CodConvenio = ? 
     AND [LisLancamentoConvenio].CodStatus = 27
     AND [LisLancamentoConvenio].DataPagamento BETWEEN ? AND ? 
     AND LOWER([LisLancamentoConvenio].HisLancamento) NOT LIKE '%estorno%' 
     and [LisLancamentoConvenio].CodRubrica = 87 order by DataPagamento"""
    
    #SEMPRE TEM Q ADICIONAR DUAS COLUNAS DEPOIS
    queryConsultaComRubricaEstorno = f"""SELECT NomeFavorecido
    ,CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-')
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     NomeTipoLancamento,   
    HisLancamento,
     NumChequeDeposito,
     DataPagamento, 
     ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio]
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado]
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado]
     WHERE 
     [LisLancamentoConvenio].CodConvenio = ? 
     AND CodStatus = 27 
     AND NomeTipoCreditoDebito = 'C' 
     AND DataPagamento BETWEEN ? AND ? 
     and [LisLancamentoConvenio].CodRubrica = 87 
     OR 
     CodStatus = 27
     AND [LisLancamentoConvenio].CodConvenio = ?  
     AND DataPagamento BETWEEN ? AND ? 
     AND LOWER(HisLancamento)  LIKE '%estorno%'
     AND [LisLancamentoConvenio].CodRubrica = 87 
     
     order by DataPagamento """


    dfconsultaDadosPorRubrica = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)
    dfconsultaDadosPorRubricaComEstorno = pd.read_sql(queryConsultaComRubricaEstorno,engine, params=parametrosComRubricaEstorno)



    return dfconsultaDadosPorRubrica,dfconsultaDadosPorRubricaComEstorno
#ok
def consultaestiloElementoDeDespesa1415Diarias(IDPROJETO,DATA1,DATA2):
    ''' Função que vai pega os dados da Rubrica 9 Despesas Financeiras e transformalos em dataframe
    para poder popular a databela Despesas Financeiras
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
    '''
    file_path = pegar_pass("passss.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)

    parametros = [(IDPROJETO, DATA1, DATA2)]
    
    queryConsultaSemEstorno = f"""
    SELECT
     [LisPessoa].[NomePessoa]
	  ,

	    CASE WHEN LEN([LisPessoa].CPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF([LisPessoa].CPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') 
        WHEN LEN([LisPessoa].CPFCNPJ) = 11 THEN STUFF(STUFF(STUFF([LisPessoa].CPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE [LisPessoa].CPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
	    (SELECT TOP 1 [NomeCidadeOrigem]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Saída') + ' - ' +  
		(SELECT TOP 1 [NomeCidadeDestino]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Saída') 

		+ ' -> ' +
		 (SELECT TOP 1 [NomeCidadeOrigem]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Chegada') + ' - ' +  
		(SELECT TOP 1 [NomeCidadeDestino]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Chegada') AS Destino


      ,[LisPagamentoDiaria].[QuantDiaria]
      ,[LisPagamentoDiaria].[ObsPedido]
      ,[LisConvenioItemAprovado].[DescConvenioItemAprovado]
	  ,[LisLancamentoConvenio].NumChequeDeposito
	  ,LisLancamentoConvenio.DataPagamento
        ,[LisLancamentoConvenio].ValorPago
    FROM [Conveniar].[dbo].[LisLancamentoConvenio]
    INNER JOIN [Conveniar].[dbo].[LisPagamentoDiaria] ON [LisLancamentoConvenio].[NumDocFinConvenio] = [LisPagamentoDiaria].[NumPedido]
    INNER JOIN [Conveniar].[dbo].[LisPessoa] ON [LisPagamentoDiaria].[CodPessoaFavorecida] = [LisPessoa].[CodPessoa]
    LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado] 
    LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado] 


    WHERE [Cod Projeto] = ? 
    AND [CodStatus] = 27
    AND LisLancamentoConvenio.DataPagamento BETWEEN ? AND ? 
    Order by LisLancamentoConvenio.DataPagamento"""

    queryConsultaComEstorno = f"""
    SELECT
     [LisPessoa].[NomePessoa]
	  ,

	    CASE WHEN LEN([LisPessoa].CPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF([LisPessoa].CPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') 
        WHEN LEN([LisPessoa].CPFCNPJ) = 11 THEN STUFF(STUFF(STUFF([LisPessoa].CPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE [LisPessoa].CPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
	    (SELECT TOP 1 [NomeCidadeOrigem]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Saída') + ' - ' +  
		(SELECT TOP 1 [NomeCidadeDestino]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Saída') 

		+ ' -> ' +
		 (SELECT TOP 1 [NomeCidadeOrigem]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Chegada') + ' - ' +  
		(SELECT TOP 1 [NomeCidadeDestino]
		FROM [Conveniar].[dbo].[LisPagamentoDiariaTrecho]
		WHERE [LisPagamentoDiariaTrecho].CodPedido = [LisPagamentoDiaria].[CodPedido]
		AND [NomeTipoDestino] = 'Chegada') AS Destino


      ,[LisPagamentoDiaria].[QuantDiaria]
      ,[LisPagamentoDiaria].[ObsPedido]
      ,[LisConvenioItemAprovado].[DescConvenioItemAprovado]
	  ,[LisLancamentoConvenio].NumChequeDeposito
	  ,LisLancamentoConvenio.DataPagamento
    ,[LisLancamentoConvenio].ValorPago
    FROM [Conveniar].[dbo].[LisLancamentoConvenio]
    INNER JOIN [Conveniar].[dbo].[LisPagamentoDiaria] ON [LisLancamentoConvenio].[NumDocFinConvenio] = [LisPagamentoDiaria].[NumPedido]
    INNER JOIN [Conveniar].[dbo].[LisPessoa] ON [LisPagamentoDiaria].[CodPessoaFavorecida] = [LisPessoa].[CodPessoa]
    LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado] 
    LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado] 


    WHERE  [LisLancamentoConvenio].CodConvenio = ?  
    AND [LisLancamentoConvenio].[CodStatus] = 27
    AND NomeTipoCreditoDebito = 'C' 
    AND LisLancamentoConvenio.DataPagamento BETWEEN ? AND ? 
    or
    [LisLancamentoConvenio].CodConvenio = ?  
    AND [LisLancamentoConvenio].[CodStatus] = 27
    AND LOWER(HisLancamento)  LIKE '%estorno%'
    AND LisLancamentoConvenio.DataPagamento BETWEEN ? AND ? 



    Order by LisLancamentoConvenio.DataPagamento"""


    dfConsultaBens = pd.read_sql(queryConsultaSemEstorno, engine, params=parametros)


    return dfConsultaBens   
#ok    
def consultaGeral30(IDPROJETO,DATA1,DATA2,codigoRubrica):

    ''' Consulta dinamica do SQL relacionado a rubrica correspondente,cada pagina tem sua própria consulta correspondente a rubrica
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
            codigoRubrica = código da rubrica 
    '''
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2,codigoRubrica)]
    parametrosComRubricaEstorno  = [(IDPROJETO, DATA1, DATA2,codigoRubrica,IDPROJETO,DATA1, DATA2,codigoRubrica,)]

    queryConsultaComRubrica = f"""SELECT NomeFavorecido,
     CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') 
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     [LisConvenioItemAprovado].[DescConvenioItemAprovado],
     
     NumDocPago,
     DataEmissao,
     NumChequeDeposito,
     DataPagamento,
     ValorPago 
     FROM [Conveniar].[dbo].[LisLancamentoConvenio] 
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado] 
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado] 
     WHERE [LisLancamentoConvenio].CodConvenio = ? 
     AND [LisLancamentoConvenio].CodStatus = 27
     AND [LisLancamentoConvenio].DataPagamento BETWEEN ? AND ? 
     AND LOWER([LisLancamentoConvenio].HisLancamento) NOT LIKE '%estorno%' 
     and [LisLancamentoConvenio].CodRubrica = ? order by DataPagamento"""
    
    queryConsultaComRubricaEstorno = f"""SELECT NomeFavorecido
    ,CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-')
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     HisLancamento,
     NumChequeDeposito,
     DataPagamento, 
     ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio]
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado]
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado]
     WHERE 
     [LisLancamentoConvenio].CodConvenio = ? 
     AND CodStatus = 27 
     AND NomeTipoCreditoDebito = 'C' 
     AND DataPagamento BETWEEN ? AND ? 
     and [LisLancamentoConvenio].CodRubrica = ? 
     OR 
     CodStatus = 27
     AND [LisLancamentoConvenio].CodConvenio = ?  
     AND DataPagamento BETWEEN ? AND ? 
     AND LOWER(HisLancamento)  LIKE '%estorno%'
     AND [LisLancamentoConvenio].CodRubrica = ? 
     
     order by DataPagamento """

    dfconsultaDadosPorRubrica = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)
    dfconsultaDadosPorRubricaComEstorno = pd.read_sql(queryConsultaComRubricaEstorno,engine, params=parametrosComRubricaEstorno)
  
    
    return dfconsultaDadosPorRubrica,dfconsultaDadosPorRubricaComEstorno

#ok
def consultaestiloElementoDeDespesa33PassagemEDespesa(IDPROJETO,DATA1,DATA2):
     
    ''' Consulta dinamica do SQL relacionado a rubrica correspondente,cada pagina tem sua própria consulta correspondente a rubrica
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
            codigoRubrica = código da rubrica 
    '''
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    parametrosComRubricaEstorno  = [(IDPROJETO, DATA1, DATA2,IDPROJETO,DATA1, DATA2,)]

    queryConsultaComRubrica = f"""SELECT NomeFavorecido,
     CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') 
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     [LisConvenioItemAprovado].[DescConvenioItemAprovado],
     
     NumDocPago,
     DataEmissao,
     NumChequeDeposito,
     DataPagamento,
     ValorPago 
     FROM [Conveniar].[dbo].[LisLancamentoConvenio] 
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado] 
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado] 
     WHERE [LisLancamentoConvenio].CodConvenio = ? 
     AND [LisLancamentoConvenio].CodStatus = 27
     AND [LisLancamentoConvenio].DataPagamento BETWEEN ? AND ? 
     AND LOWER([LisLancamentoConvenio].HisLancamento) NOT LIKE '%estorno%' 
     and [LisLancamentoConvenio].CodRubrica in (20,78,52)  order by [LisLancamentoConvenio].DataPagamento"""
    
    queryConsultaComRubricaEstorno = f"""SELECT NomeFavorecido
    ,CASE WHEN LEN(FavorecidoCPFCNPJ) > 11 THEN STUFF(STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-')
     WHEN LEN(FavorecidoCPFCNPJ) = 11 THEN STUFF(STUFF(STUFF(FavorecidoCPFCNPJ, 4, 0, '.'), 8, 0, '.'), 12, 0, '-') ELSE FavorecidoCPFCNPJ END AS FormattedFavorecidoCPFCNPJ,
     HisLancamento,
     NumChequeDeposito,
     DataPagamento, 
     ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio]
     LEFT JOIN [Conveniar].[dbo].[PlanoTrabalhoLancamento] ON [LisLancamentoConvenio].[CodLancamento] = [PlanoTrabalhoLancamento].[CodLancamentoGerado]
     LEFT JOIN [Conveniar].[dbo].[LisConvenioItemAprovado] ON [PlanoTrabalhoLancamento].[CodConvenioItemAprovado] = [LisConvenioItemAprovado].[CodConvenioItemAprovado]
     WHERE 
     [LisLancamentoConvenio].CodConvenio = ? 
     AND CodStatus = 27 
     AND NomeTipoCreditoDebito = 'C' 
     AND DataPagamento BETWEEN ? AND ? 
     and [LisLancamentoConvenio].CodRubrica in (20,78,52) 
     OR 
     CodStatus = 27
     AND [LisLancamentoConvenio].CodConvenio = ?  
     AND DataPagamento BETWEEN ? AND ? 
     AND LOWER(HisLancamento)  LIKE '%estorno%'
     AND [LisLancamentoConvenio].CodRubrica in (20,78,52)
     
     order by [LisLancamentoConvenio].DataPagamento """

    dfconsultaDadosPorRubrica = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)
    dfconsultaDadosPorRubricaComEstorno = pd.read_sql(queryConsultaComRubricaEstorno,engine, params=parametrosComRubricaEstorno)
  
    
    return dfconsultaDadosPorRubrica,dfconsultaDadosPorRubricaComEstorno

#ok
def consultaBens(IDPROJETO,DATA1,DATA2):
    ''' Função que vai pega os dados da Rubrica 9 Despesas Financeiras e transformalos em dataframe
    para poder popular a databela Despesas Financeiras
        Argumentos
            IDPROJETO = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
    '''
    file_path = pegar_pass("passss.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    idprojetoComZero = f"0{IDPROJETO}"
    parametros = [(IDPROJETO,idprojetoComZero, DATA1, DATA2)]
    
    queryConsultaComRubrica = f"""
    SELECT [Descrição][descri],
    [Patrimônio][patri],
    [Data de Aquisição][dataAqui],
    [Nº Nota][nota],
    [Localização][localiza]
   ,[Descrição],
    [Valor de Aquisição][valorAqui],
    [Valor de Aquisição][valorAqui2],
    [Responsável][responsavel] 
    FROM [SBO_FINATEC].[dbo].[VW_BENS_ADQUIRIDOS] 
    WHERE ([Cod Projeto] = ? or [Cod Projeto] = ? ) 
    AND [Status] = 'Imobilizado' 
    AND [Data de Aquisição] BETWEEN ? AND ? 
    Order by [Data de Aquisição]"""
    dfConsultaBens = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)

    print(dfConsultaBens)
    return dfConsultaBens   
#ok    
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
    parametros2 = [(IDPROJETO, DATA1, DATA2, IDPROJETO, DATA1, DATA2)]
    #consultaSemEstorno = f"SELECT DISTINCT DataPagamento,ValorPago,NumChequeDeposito,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) NOT LIKE '%estorno%' order by DataPagamento"
    #consultaComEstorno =  f"SELECT DISTINCT DataPagamento,ValorPago,NumChequeDeposito,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) LIKE '%estorno%' order by DataPagamento"
    
    #consultaloucona assim por que precisa ser detalhado cada item na tabela
    consultaSemEstorno = f"""
    SELECT 
  	[LisLancamentoConvenio].HisLancamento, 
	CONVERT(varchar, CAST([LisLancamentoConvenio].DataPagamento AS datetime), 103) AS FormattedDate,
    [LisPagamentoDespesaConvenioAdministrativa].Valor 
    FROM [Conveniar].[dbo].[LisLancamentoConvenio]
    INNER JOIN  [Conveniar].[dbo].[LisDocumentoConvenio] ON [LisLancamentoConvenio].[CodDocFinConvenio] = [LisDocumentoConvenio].[CodDocFinConvenio]
    INNER JOIN  [Conveniar].[dbo].[DocFinConvPagDespesa] ON [LisDocumentoConvenio].[CodDocFinConvenio] = [DocFinConvPagDespesa].[CodDocFinConvenio]
    INNER JOIN  [Conveniar].[dbo].[LisPagamentoDespesaConvenio] ON [DocFinConvPagDespesa].[CodPedido] = [LisPagamentoDespesaConvenio].[CodPedido]
    INNER JOIN  [Conveniar].[dbo].[LisPagamentoDespesaConvenioAdministrativa] ON [LisPagamentoDespesaConvenio].CodDespesaConvenio = [LisPagamentoDespesaConvenioAdministrativa].CodDespesaConvenio
    AND [LisPagamentoDespesaConvenio].CodConvenio = [LisPagamentoDespesaConvenioAdministrativa].CodConvenio
    WHERE [LisLancamentoConvenio].CodConvenio = ? AND [LisLancamentoConvenio].CodStatus = 27 AND [LisLancamentoConvenio].CodRubrica = 9 AND [LisLancamentoConvenio].DataPagamento BETWEEN ? AND ?
    AND LOWER([LisLancamentoConvenio].HisLancamento) NOT LIKE '%estorno%' order by [LisLancamentoConvenio].DataPagamento"""

    consultaComEstorno = f"""SELECT 
    HisLancamento ,
    CONVERT(varchar, CAST(DataPagamento AS datetime), 103) AS FormattedDate,
    ValorPago
    
    FROM [Conveniar].[dbo].[LisLancamentoConvenio] 
    WHERE CodConvenio = ? 
    AND CodStatus = 27 
    AND CodRubrica = 9 
    AND DataPagamento BETWEEN ? AND ? 
    AND LOWER(HisLancamento)  LIKE '%estorno%'
    OR
    CodConvenio = ? 
    AND CodStatus = 27 
    AND CodRubrica = 9 
    AND NomeTipoCreditoDebito = 'C' 
    AND DataPagamento BETWEEN  ? AND ? 
    order by DataPagamento"""
    
    dfSemEstorno = pd.read_sql(consultaSemEstorno, engine, params=parametros)
    dfComEstorno = pd.read_sql(consultaComEstorno, engine, params=parametros2)
    

    return dfSemEstorno,dfComEstorno
#ok
def consultaRendimentosAplicacao(IDPROJETO,DATA1,DATA2):
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2)]
    consultaRendimentoAplicacao = f"SELECT  DataPagamento,ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 3 and NomeTipoLancamento = 'Aplicação Financeira'  AND DataPagamento BETWEEN ? AND ?  order by DataPagamento"
    dfConsultaRendimentoAplicacao = pd.read_sql(consultaRendimentoAplicacao, engine, params=parametros)
     
    consultaRendimentoEImposto =  f"SELECT DataPagamento,ValorPago,NomeTipoLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 3 and (NomeTipoLancamento = 'Aplicação Financeira' or NomeTipoLancamento = 'IRRF Pessoa Jurídica')  AND DataPagamento BETWEEN ? AND ?  order by DataPagamento"
    dfConsultaRendimentoEImposto = pd.read_sql(consultaRendimentoEImposto, engine, params=parametros)
    
    consultaImposto =  f"SELECT  DataPagamento,ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 3 and NomeTipoLancamento = 'IRRF Pessoa Jurídica'  AND DataPagamento BETWEEN ? AND ?  order by DataPagamento"
    dfConsultaImposto = pd.read_sql(consultaImposto, engine, params=parametros)
     
   
    return dfConsultaRendimentoAplicacao,dfConsultaImposto,dfConsultaRendimentoEImposto




#preencher


def preencheFinep(codigo,data1,data2,tabela):
    '''Preenche a planilha fap

        Argumentos: 
            codigo = CodConvenio na tabela nova, corresponde ao codigo do projeto
            DATA1 = Data Inicial Selecinado pelo Usuario
            DATA2 = Data Final Selecionado pelo Usuario
            tabela = tabela a ser preenchida  extensão xlsx
   '''
    consultaRelatorioExecFinanceiraA1(codigo,data1,data2)
    consultaDemonstrativoReceitaEDespesaA2(codigo,data1,data2)
    consultaPagamentoPessoal(codigo,data1,data2)
    # consultaestiloElementoDeDespesa1415Diarias(codigo,data1,data2)
    # consultaGeral30(codigo,data1,data2,87)
    # consultaestiloElementoDeDespesa33Diarias(codigo,data1,data2)
    consultaBens(codigo,data1,data2)
    consultaConciliacaoBancaria(codigo,data1,data2)
    consultaRendimentosAplicacao(codigo,data1,data2)