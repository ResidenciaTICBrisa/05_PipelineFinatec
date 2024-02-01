import pyodbc
from datetime import datetime,date
import openpyxl
import os
from .estilo_fub import *
from collections import defaultdict
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

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
    consultaSemEstorno = f"SELECT DISTINCT DataPagamento,ValorLancamento,NumChequeDeposito,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) NOT LIKE '%estorno%' order by DataPagamento"
    consultaComEstorno =  f"SELECT DISTINCT DataPagamento,ValorLancamento,NumChequeDeposito,HisLancamento FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND CodRubrica = 9 AND DataPagamento BETWEEN ? AND ? AND LOWER(HisLancamento) LIKE '%estorno%' order by DataPagamento"
    dfSemEstorno = pd.read_sql(consultaSemEstorno, engine, params=parametros)
    dfComEstorno = pd.read_sql(consultaComEstorno, engine, params=parametros)
   

    return dfSemEstorno,dfComEstorno

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


    values = [d.get(key) for d in list_of_dicts for key in keys]
    
    #print(values)
    return values

def consultaNomeRubricaCodRubrica(IDPROJETO, DATA1, DATA2,):
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
    queryNomeRubrica = f"SELECT CodRubrica,NomeRubrica FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? GROUP BY NomeRubrica, CodRubrica"
    dfNomeRubricaCodRubrica = pd.read_sql(queryNomeRubrica, engine, params=parametros)


    return dfNomeRubricaCodRubrica

def consultaProjeto(IDPROJETO, DATA1, DATA2,codigoRubrica):
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
    parametros = [(IDPROJETO, DATA1, DATA2,codigoRubrica)]
    queryConsultaComRubrica = f"SELECT NomeFavorecido,FavorecidoCPFCNPJ,NomeTipoLancamento,HisLancamento,NumDocPago,DataEmissao,NumChequeDeposito,DataPagamento, ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? and CodRubrica = ? "
    dfconsultaDadosPorRubrica = pd.read_sql(queryConsultaComRubrica, engine, params=parametros)


    return dfconsultaDadosPorRubrica

def consultaEntradaReceitas(IDPROJETO, DATA1, DATA2):
  
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    parametros = [(IDPROJETO, DATA1, DATA2,)]
    queryConsultaReceita = f"SELECT DataPagamento,NumChequeDeposito,NomeFavorecido,ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND  (CodRubrica = 2 OR CodRubrica = 67 OR CodRubrica = 88) AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? ORDER BY  ValorPago  DESC"
    queryConsultaDemonstrativoReceita = f"SELECT DataPagamento,HisLancamento,NumChequeDeposito,ValorPago FROM [Conveniar].[dbo].[LisLancamentoConvenio] WHERE CodConvenio = ? AND (CodRubrica = 2 OR CodRubrica = 67 OR CodRubrica = 88) AND CodStatus = 27 AND DataPagamento BETWEEN ? AND ? ORDER BY  ValorPago DESC"
    dfReceitas = pd.read_sql(queryConsultaReceita, engine, params=parametros)
    dfDemonstrativoReceitas = pd.read_sql(queryConsultaDemonstrativoReceita, engine, params=parametros)


    return dfReceitas,dfDemonstrativoReceitas

def consultaReceitaEExecReceita(IDPROJETO, DATA1, DATA2):
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
    
    dfComPeriodo= pd.read_sql(consultaComPeriodo, engine, params=parametros)
    dfAteAData = pd.read_sql(consultaAteAData, engine, params=parametros2)
    dfPrevisto = pd.read_sql(consultaPrevisto, engine, params=(IDPROJETO,))
   

    return dfComPeriodo,dfAteAData,dfPrevisto    

#preenche planilha 

def conciliacaoBancaria(codigo,data1,data2,planilha,stringTamanho):

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
                # print(row_num)
                # print(col_num)
       
        linha2 = 16+4+tamanho


        for row_num, row_data in enumerate(dataframeComEstorno.itertuples(index=False), start=linha2):#inicio linha
            for col_num, value in enumerate(row_data, start=1):#inicio coluna
                worksheet333.cell(row=row_num, column=col_num, value=value)
                # print("comestorno")
                # print(row_num)
                # print(col_num)
        

        workb.save(tabela)
        workb.close

def demonstrativo(codigo,data1,data2,planilha,rowBrasilia,dataframeDemonstrativoReceita):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Demonstrativo de Receita")
    workbook.save(tabela)
    workbook.close()
    tamanho = len(dataframeDemonstrativoReceita)
    estilo_demonstrativoDeReceita(tabela,tamanho,rowBrasilia)

    workbook = openpyxl.load_workbook(tabela)
    worksheet = workbook['Demonstrativo de Receita']
   

    for row_num, row_data in enumerate(dataframeDemonstrativoReceita.itertuples(index = False), start=10):#inicio linha
        for col_num, value in enumerate(row_data, start=1):#inicio coluna
    
            worksheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))
    
    workbook.save(tabela)
    workbook.close()

def relacaodeBens(codigo,data1,data2,planilha,rowBrasilia):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    nomeTabela ="Relação de Bens"
    tituloStyle = "relacaoBEns"
    sheet2 = workbook.create_sheet(title="Relação de Bens")
    workbook.save(tabela)
    workbook.close()
    tamanho = 20
    estiloRelacaoBens(tabela,tamanho,tituloStyle,nomeTabela,rowBrasilia)

def rendimentoDeAplicacao(codigo,data1,data2,planilha,rowBrasilia):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Rendimento de Aplicação")
    workbook.save(tabela)
    workbook.close()
    tamanho = 20
    estilo_rendimento_de_aplicacao(tabela,tamanho,rowBrasilia) 

def rubricaGeral(codigo,data1,data2,planilha,rowBrasilia):
     #consulta todas rubricas
    tabela = pegar_caminho(planilha)
    dfNomeRubricaCodigoRubrica = consultaNomeRubricaCodRubrica(codigo, data1, data2)

    for index, values in dfNomeRubricaCodigoRubrica.iterrows():
        
        dfConsultaProjeto = consultaProjeto(codigo, data1, data2,values['CodRubrica'])
        # print(dfConsultaProjeto)
        if values['NomeRubrica'] == "Obrigações Tributárias e contributivas":
            values['NomeRubrica'] == "Obrigações Tributárias"
        if values['NomeRubrica'] == f"Obrigações Tributárias e Contributivas - 20% de OST" :
            values['NomeRubrica'] == f"Obrigações Trib. - Encargos 20%"
        if values['NomeRubrica'] == f"Outros Serviços de Terceiros - Pessoa Física" :
            values['NomeRubrica'] == f"Outros Serviços Terceiros - PF"
        if values['NomeRubrica'] == f"Outros Serviços de Terceiros - Pessoa Jurídica" :
            values['NomeRubrica'] == f"Outros Serviços Terceiros - PJ"
        if values['NomeRubrica'] == f"Passagens e Despesas com Locomoção" :
            values['NomeRubrica'] == f"Passagens e Desp. Locomoção"
        

        
        if values['NomeRubrica'] != "Rendimentos de Aplicações Financeiras" and values['NomeRubrica'] != "Despesas Financeiras" and values['NomeRubrica'] != "Receitas":
            nomeTabela = values['NomeRubrica']
            tituloStyle = values['NomeRubrica']
            workbook = openpyxl.load_workbook(tabela)
            sheet2 = workbook.create_sheet(title=f"{values['NomeRubrica']}")
            workbook.save(tabela)
            workbook.close()

            tamanho = len(dfConsultaProjeto)
            # print("tamanhodics")
            # print(tamanho)
            estiloGeral(tabela,tamanho,tituloStyle,nomeTabela,rowBrasilia)
            workbook = openpyxl.load_workbook(tabela)
            sheet2 = workbook[values['NomeRubrica']]
            for row_num, row_data in enumerate(dfConsultaProjeto.itertuples(), start=10):#inicio linha
                for col_num, value in enumerate(row_data, start=1):#inicio coluna
                    value = convert_datetime_to_stringdt(value)
                    sheet2.cell(row=row_num, column=col_num, value=value)
                        # print(row_num)
                    # print(col_num)
                    # print(value)
            
            workbook.save(tabela)
            workbook.close()

       

     #consulta individual
    return 0

def Receita(planilha,codigo,data1,data2,tamanhoResumo,dataframe):
    caminho = pegar_caminho(planilha)
    if len(dataframe) > tamanhoResumo :
        tamanhoResumo  = len(dataframe)
    tamanho,tamanhoequipamentos = estiloReceitaXDespesa(caminho,tamanhoResumo)
    if len(dataframe) > tamanhoResumo :
        tamanho  = len(dataframe)
    #o tamanho na verdade tem que ser do data frame se ele for maior
   
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
    print(tamanhoequipamentos)
    print(tamanhoequipamentos)
     #Obras e Instalações
    #previsto
    string_exists = dataframe['NomeRubrica'].isin(["Obras e Instalações"]).any()
    if string_exists:
        
        #periodo
        stringObras = f'I{tamanhoequipamentos -1}'
        sheet[stringObras] = dataframe.loc[dataframe['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_PERIODO'].values[0]
        

    string_exists = dataframe['NomeRubrica'].isin(["Equipamentos e Material Permanente"]).any()
    if string_exists:
    #Materiais Equipamentos e Material Permanente   
        
        #periodo
        stringObras = f'I{tamanhoequipamentos}'
        sheet[stringObras] = dataframe.loc[dataframe['NomeRubrica'] == 'Equipamentos e Material Permanente', 'VALOR_TOTAL_PERIODO'].values[0]

    #Materiais Equipamentos e Material nACIONAL
    string_exists = dataframe['NomeRubrica'].isin(["Material Permanente e Equipamento Nacional"]).any()
    if string_exists:
      
        #periodo
        stringObras = f'I{tamanhoequipamentos +1}'
        sheet[stringObras] = dataframe.loc[dataframe['NomeRubrica'] == 'Material Permanente e Equipamento Nacional', 'VALOR_TOTAL_PERIODO'].values[0]
       

    #Materiais Equipamentos e Material iMPORTADO
    string_exists = dataframe['NomeRubrica'].isin(["Material Permanente e Equipamento Importado"]).any()
    if string_exists:
      
        #periodo
        stringObras = f'I{tamanhoequipamentos + 2}'
        sheet[stringObras] = dataframe.loc[dataframe['NomeRubrica'] == 'Material Permanente e Equipamento Importado', 'VALOR_TOTAL_PERIODO'].values[0]
       
    print(tamanhoequipamentos+3)
    string_exists = dataframe['NomeRubrica'].isin(["Rendimentos de Aplicações Financeiras"]).any()
    if string_exists:
        #periodo
        stringObras = f'I{tamanhoequipamentos + 3}'
        sheet[stringObras] = sheet[stringObras] = dataframe.loc[dataframe['NomeRubrica'] == 'Rendimentos de Aplicações Financeiras', 'VALOR_TOTAL_PERIODO'].values[0]
       
      

    values_to_remove = ["Receitas", "Rendimentos de Aplicações Financeiras", "Despesas Financeiras",'Material Permanente e Equipamento Nacional','Material Permanente e Equipamento Importado']
    dataframe = dataframe[~dataframe['NomeRubrica'].isin(values_to_remove)]
    for row_num, row_data in enumerate(dataframe.itertuples(index = False), start=16):#inicio linha
        for col_num, value in enumerate(row_data, start=8):#inicio coluna
                sheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))


    dfReceitas,dfDemonstrativoReceitas = consultaEntradaReceitas(codigo,data1,data2)


    for row_num, row_data in enumerate(dfReceitas.itertuples(index = False), start=16):#inicio linha
        for col_num, value in enumerate(row_data, start=1):#inicio coluna
                if col_num == 4:
                     col_num = 5
                sheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))

              

    workbook.save(planilha)
    workbook.close()
    print(strintT)
    return strintT,dfDemonstrativoReceitas

def ExeReceitaDespesa(planilha,codigo,data1,data2,stringTamanho):
    tabela = pegar_caminho(planilha)
    workbook = openpyxl.load_workbook(tabela)
    sheet2 = workbook.create_sheet(title="Exec. Receita e Despesa")
    workbook.save(tabela)
    workbook.close()
    stringTamanho = 15 #esqueci o pq deve ser  =o tamanho na tabela

    #dataframe com os dados
    dfComPeriodo,dfAteAData,dfPrevisto = consultaReceitaEExecReceita(codigo,data1,data2)
    # Merge with an outer join
    merged_df = pd.merge(dfPrevisto, dfComPeriodo, on='NomeRubrica', how='outer')
    dfMerged = pd.merge(merged_df,dfAteAData, on = 'NomeRubrica', how = 'outer')
    tamanho = len(dfMerged)#tamanho para deixar dinamico para imprimir sa rubricas
    string_exists = dfMerged['NomeRubrica'].isin(["Material Permanente e Equipamento Importado"]).any()
    if string_exists:
         tamanho = tamanho - 1
    string_exists = dfMerged['NomeRubrica'].isin(["Obras e Instalações"]).any()
    if string_exists:
         tamanho = tamanho - 1
    string_exists = dfMerged['NomeRubrica'].isin(["Equipamentos e Material Permanente"]).any()
    if string_exists:
         tamanho = tamanho - 1
    string_exists = dfMerged['NomeRubrica'].isin(["Material Permanente e Equipamento Nacional"]).any()
    if string_exists:
         tamanho = tamanho - 1
    
    tamanho = tamanho - 3

    stringTamanho = tamanho + 16 
    estiloExecReceitaDespesa(tabela,tamanho,stringTamanho)
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

    #Obras e Instalações
    #previsto
    string_exists = dfMerged['NomeRubrica'].isin(["Obras e Instalações"]).any()
    if string_exists:
        
        stringObras = f'B{stringTamanho + 2}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_PREVISTO'].values[0]
        stringObras = f'F{stringTamanho + 2}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_PREVISTO'].values[0]

        #periodo
        stringObras = f'C{stringTamanho + 2}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_PERIODO'].values[0]
        #Ate o momento
        stringObras = f'G{stringTamanho + 2}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_DATA'].values[0]

    string_exists = dfMerged['NomeRubrica'].isin(["Equipamentos e Material Permanente"]).any()
    if string_exists:
    #Materiais Equipamentos e Material Permanente   
        stringObras = f'B{stringTamanho + 3}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Equipamentos e Material Permanente', 'VALOR_TOTAL_PREVISTO'].values[0]
        stringObras = f'F{stringTamanho + 3}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Equipamentos e Material Permanente', 'VALOR_TOTAL_PREVISTO'].values[0]
        #periodo
        stringObras = f'C{stringTamanho + 3}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Equipamentos e Material Permanente', 'VALOR_TOTAL_PERIODO'].values[0]
        #Ate o momento
        stringObras = f'G{stringTamanho + 3}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Obras e Instalações', 'VALOR_TOTAL_DATA'].values[0]

    #Materiais Equipamentos e Material nACIONAL
    string_exists = dfMerged['NomeRubrica'].isin(["Material Permanente e Equipamento Nacional"]).any()
    if string_exists:
        stringObras = f'B{stringTamanho + 4}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Nacional', 'VALOR_TOTAL_PREVISTO'].values[0]
        stringObras = f'F{stringTamanho + 4}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Nacional', 'VALOR_TOTAL_PREVISTO'].values[0]
        #periodo
        stringObras = f'C{stringTamanho + 4}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Nacional', 'VALOR_TOTAL_PERIODO'].values[0]
        #Ate o momento
        stringObras = f'G{stringTamanho + 4}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Nacional', 'VALOR_TOTAL_DATA'].values[0]

    #Materiais Equipamentos e Material iMPORTADO
    string_exists = dfMerged['NomeRubrica'].isin(["Material Permanente e Equipamento Importado"]).any()
    if string_exists:
        stringObras = f'B{stringTamanho + 5}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Importado', 'VALOR_TOTAL_PREVISTO'].values[0]
        stringObras = f'F{stringTamanho + 5}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Importado', 'VALOR_TOTAL_PREVISTO'].values[0]
        #periodo
        stringObras = f'C{stringTamanho + 5}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Importado', 'VALOR_TOTAL_PERIODO'].values[0]
        #Ate o momento
        stringObras = f'G{stringTamanho + 5}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Material Permanente e Equipamento Importado', 'VALOR_TOTAL_DATA'].values[0]

    string_exists = dfMerged['NomeRubrica'].isin(["Rendimentos de Aplicações Financeiras"]).any()
    if string_exists:
        stringObras = f'B{stringTamanho + 7}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Rendimentos de Aplicações Financeiras', 'VALOR_TOTAL_PREVISTO'].values[0]
        stringObras = f'F{stringTamanho + 7}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Rendimentos de Aplicações Financeiras', 'VALOR_TOTAL_PREVISTO'].values[0]
        #periodo
        stringObras = f'C{stringTamanho + 7}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Rendimentos de Aplicações Financeiras', 'VALOR_TOTAL_PERIODO'].values[0]
        #Ate o momento
        stringObras = f'G{stringTamanho + 7}'
        sheet[stringObras] = dfMerged.loc[dfMerged['NomeRubrica'] == 'Rendimentos de Aplicações Financeiras', 'VALOR_TOTAL_DATA'].values[0]


   

    # print(dfMerged)

    #remover essas linhas da tabela
    values_to_remove = ["Receitas", "Rendimentos de Aplicações Financeiras", "Despesas Financeiras",'Material Permanente e Equipamento Nacional','Material Permanente e Equipamento Importado']

    # Use boolean indexing to drop rows based on the values in the first column
    dfMerged = dfMerged[~dfMerged['NomeRubrica'].isin(values_to_remove)]
    # print(dfMerged)
    for row_num, row_data in enumerate(dfMerged.itertuples(index = False), start=16):#inicio linha
        for col_num, value in enumerate(row_data, start=1):#inicio coluna
                if col_num == 2:
                    sheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))
                    col_num = 6
                    sheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))
                    col_num=2

                if col_num == 4:
                    col_num = 7
                sheet.cell(row=row_num, column=col_num, value=convert_datetime_to_string(value))
                    # print(row_num)
                    # print(col_num)
                    # print(value)
            

            


    workbook.save(planilha)
    workbook.close()
    return tamanho,dfComPeriodo

def preencheFub(codigo,data1,data2,tabela):
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


    '''
    tamanho,dataframe = ExeReceitaDespesa(tabela,codigo,data1,data2,15)
    tamanhoPosicaoBrasilia,dataframeDemonstrativoReceita = Receita(tabela,codigo,data1,data2,tamanho,dataframe)
    demonstrativo(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia,dataframeDemonstrativoReceita)
    rubricaGeral(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    conciliacaoBancaria(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    rendimentoDeAplicacao(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    relacaodeBens(codigo,data1,data2,tabela,tamanhoPosicaoBrasilia)
    
