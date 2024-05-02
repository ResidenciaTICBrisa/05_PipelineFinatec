import pyodbc
import os

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



file_path = pegar_pass("passs.txt")
conStr = ''
with open(file_path, 'r') as file:
            conStr = file.readline().strip()

conn = pyodbc.connect(conStr)
cursor = conn.cursor()

print(cursor)

stringconsulta = """SELECT TOP 1000 [Pedido].[CodPedido]
      ,[Pedido].[NumPedido]
 ,[ArquivoBinario].[ArquivoBinario]
 ,[Arquivo].*
 ,[ArquivoBinario].*

  FROM [Conveniar].[dbo].[Pedido]

  LEFT JOIN [Conveniar].[dbo].[Arquivo] ON [Pedido].[CodPedido] = [Arquivo].[CodSolicitacao]
  LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoReferencia] ON [Arquivo].[CodArquivoReferencia] = [ArquivoReferencia].CodArquivoReferencia
  LEFT JOIN [ConveniarArquivo].[dbo].[ArquivoBinario] ON [ArquivoReferencia].ChaveLocalArmazenamento = [ArquivoBinario].[CodArquivoBinario]

  WHERE NumPedido = '233412024'"""
with open("Output.pdf", "wb") as output_file:
    cursor.execute(stringconsulta)
    ablob = cursor.fetchone()
    output_file.write(ablob[2])

cursor.close()
conn.close()