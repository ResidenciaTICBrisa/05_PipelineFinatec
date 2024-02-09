import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import os


def pegar_pass(chave):
    arq_atual = os.path.abspath(__file__)
    app = os.path.dirname(arq_atual)
    project = os.path.dirname(app)
    pipeline = os.path.dirname(project)
    desktop = os.path.dirname(pipeline)
    caminho_pipeline = os.path.join(desktop, chave)
    
    return caminho_pipeline

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
    print(engine)
    idprojetoComZero = f"0{IDPROJETO}"
    #parametros = [(IDPROJETO,idprojetoComZero, DATA1, DATA2)]
    
    #queryConsultaComRubrica = f"SELECT [Descrição][descri],[Patrimônio][patri],[Data de Aquisição][dataAqui],[Nº Nota][nota],[Localização][localiza],[telefone],[Valor de Aquisição][valorAqui],[Valor de Aquisição][valorAqui2],[Responsável][responsavel] FROM [SBO_FINATEC].[dbo].[VW_BENS_ADQUIRIDOS] WHERE ([Cod Projeto] = ? or [Cod Projeto] = ? ) AND [Status] = 'Imobilizado' AND [Data de Aquisição] BETWEEN ? AND ? Order by [Data de Aquisição]"
    queryConsultaComRubrica = f"SELECT *  FROM [SBO_FINATEC].[dbo].[VW_BENS_ADQUIRIDOS]"
    dfConsultaBens = pd.read_sql(queryConsultaComRubrica, engine)

    return dfConsultaBens





consultaBens(6585,'2022-01-01 00:00:00','2023-01-31 00:00:00')