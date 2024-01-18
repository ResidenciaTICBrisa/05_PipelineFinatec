import oracledb
import os
import datetime


def convert_datetime_to_string(value):  
    if isinstance(value, datetime.datetime):
        return value.strftime('%d/%m/%Y')
    return value
#connection string in the format
#<username>/<password>@<dBhostAddress>:<dbPort>/<dbServiceName>


def getCollumNames(IDPROJETO):
# def getCollumNames(IDPROJETO, DATA1, DATA2):
    file_path = "/home/ubuntu/Desktop/devfront/devfull/pass.txt"
    conStr = ''
    with open(file_path, 'r') as file:
            conStr = file.readline().strip()

    conn = None
    conn = oracledb.connect(conStr)
    cur = conn.cursor()

    sql = """SELECT DISTINCT * FROM IDEA.FAT_LANCAMENTO_CONVENIAR 
             WHERE ID_PROJETO = :IDPROJETO 
             AND ID_STATUS = 27  
             ORDER BY NUM_DOC_FIN"""
    # sql = """SELECT DISTINCT * FROM IDEA.FAT_LANCAMENTO_CONVENIAR 
    #          WHERE ID_PROJETO = :IDPROJETO 
    #          AND ID_STATUS = 27 
    #          AND DATA_PAGAMENTO BETWEEN TO_DATE(:DATA1, 'YYYY-MM-DD') 
    #          AND TO_DATE(:DATA2, 'YYYY-MM-DD') 
    #          ORDER BY NUM_DOC_FIN"""

    cur.execute(sql, {
        'IDPROJETO': IDPROJETO
    })
    # cur.execute(sql, {
    #     'IDPROJETO': IDPROJETO,
    #     'DATA1': DATA1,
    #     'DATA2': DATA2
    # })
    return cur



def get_values_from_dict(codigo):
  
    gete = getCollumNames(codigo)

    collums = []
    for i in gete.description:
        collums.append(i[0])
    
    print(collums)

    value = []
    for i in gete:
        val = tuple(convert_datetime_to_string(item) for item in i)
        value.append(val)
    #print(value)
    list_of_dicts = [dict(zip(collums, values)) for values in value]

    #print(list_of_dicts)
    return list_of_dicts

def retornavalores(list_of_dicts,keys):
    values = [d.get(key) for d in list_of_dicts for key in keys]
    
    #print(values)
    return values

valor = get_values_from_dict(7262)
# print(type(valor))

# # Step 1: Extract unique values from the 'ID_RUBRICA' key
unique_id_rubrica_values = set(item['ID_RUBRICA'] for item in valor)

# # Step 2: Create separate lists of dictionaries for each unique 'ID_RUBRICA' value
categorized_data = {value: [] for value in unique_id_rubrica_values}
for item in valor:
    categorized_data[item['ID_RUBRICA']].append(item)

#print(categorized_data)

categoriz = str(categorized_data)
with open('saves.txt', 'w') as file:
    # Writing the variable to the file
    file.write(categoriz)