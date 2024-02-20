import pyodbc
import os
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL


def pegar_pass(chave):
    arq_atual = os.path.abspath(__file__)
    app = os.path.dirname(arq_atual)
    project = os.path.dirname(app)
    pipeline = os.path.dirname(project)
    desktop = os.path.dirname(pipeline)
    caminho_pipeline = os.path.join(desktop, chave)
    
    return caminho_pipeline


    # return records


def consultaTudo():
    
    file_path = pegar_pass("passs.txt")
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conStr})
    engine = create_engine(connection_url)
    queryConsultaComRubrica = f"SELECT * FROM [Conveniar].[dbo].[LisConvenio]"
    dfconsultaDadosPorRubrica = pd.read_sql(queryConsultaComRubrica, engine)

    return dfconsultaDadosPorRubrica

def consultaQtd(df, init, end):
    final = init + end
    slice_df = df.iloc[init:final, :]

    return slice_df

def consultaCodConvenio(df):
    return df['CodConvenio'].tolist()

def consultaLimitedDict(df, length):
    print("")



