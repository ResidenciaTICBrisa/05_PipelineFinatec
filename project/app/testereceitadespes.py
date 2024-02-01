import pandas
import pyodbc
import os


    
def pegar_pass(chave):
    arq_atual = os.path.abspath(__file__)
    app = os.path.dirname(arq_atual)
    project = os.path.dirname(app)
    pipeline = os.path.dirname(project)
    desktop = os.path.dirname(pipeline)
    caminho_pipeline = os.path.join(desktop, chave)
    
    return caminho_pipeline

def retornavalores(list_of_dicts,keys):
    values = [d.get(key) for d in list_of_dicts for key in keys]
    
    #print(values)
    return values

def queryReceitaXDespesa(CodConvenio,DATA1,DATA2,rubricaprincipal,rubricas,rubricaprincipalstring):
    
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
            print(valores_preenchimento)
            for i in range(len(valores_preenchimento)):
                soma = soma + valores_preenchimento[i]
    
    dicionariosaida[rubricaprincipalstring] = soma

    cursor.close()
    conn.close()

    return dicionariosaida

#pessoal fisica
rubricas = [79,54,55]
dict1 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',25,rubricas,'Outros Serviços de Terceiros - Pessoa Física')

#pessoa Juridica
rubricas = [57,26]
dict2 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',75,rubricas,'Serviços de Terceiros Pessoa Jurídica')

#iss2
dict3 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',88,rubricas,'Encargos - ISS 2% ')

#ISS5
dict4 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',67,rubricas,'Encwerwerwerwewargos - ISS 5% ')

 #passagemLocomoção
rubricas = [52,20]
dict5 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',7,rubricas,'Encrtertargos - ISS 5% ')
 
 #serv.terceiro celetista
rubricas = [81327,68132,61239,7013]
dict6 = queryReceitaXDespesa(6477,'2012-01-01 00:00:00','2023-01-31 00:00:00',7132,rubricas,'0000000 - ISS 5% ')


merged_dict = {**dict1, **dict2, **dict3, **dict4, **dict5,**dict6}
print("so o dicionario")
non_zero_dict = {}
for key, value in merged_dict.items() : 
        if value != 0:
            non_zero_dict[key] = value

#print(non_zero_dict)
print(len(non_zero_dict))

for key,value in non_zero_dict.items():
     print(f'Chave{key}')
     print(f'valor{value}')


