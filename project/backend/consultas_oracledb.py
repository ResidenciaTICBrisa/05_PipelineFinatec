import oracledb


file_path = "/home/ubuntu/Desktop/devfront/devfull/pass.txt"
conStr = ''
with open(file_path, 'r') as file:
        conStr = file.readline().strip()

def getCollumNames():

    #inicializando o objeto que ira conectar no db
    conn = None
    #criando o objeto de conexão das
    conn = oracledb.connect(conStr)
    #criar um objeto cursor necessario para fazer as consultas
    cur = conn.cursor() 
    cur.execute("SELECT * FROM IDEA.STG_PROJETOS_CONVENIAR")
    # colunas_desejadas = [
    #     'CODIGO', 'NOME', 'NOME_FINANCIADOR', 
    #     'DATA_ASSINATURA', 'DATA_VIGENCIA', 'COORDENADOR', 
    #     'VALOR_APROVADO', 'GRUPO_GESTORES'
    # ]

    # # Montar a consulta SQL selecionando apenas as colunas desejadas
    # consulta_sql = "SELECT {} FROM IDEA.STG_PROJETOS_CONVENIAR".format(", ".join(colunas_desejadas))
    

    return cur

print("\n")

# cur.close()
# #encerra a conexao
# conn.close()
# print("conexão db completa!")

def getlimitedRows(numb):
    consulta = {}
    a=[]
    try:
        connection = oracledb.connect(conStr)
        cursor = connection.cursor()
        print("Connected to database")
        sqlite_select_query = f"SELECT * FROM IDEA.STG_PROJETOS_CONVENIAR WHERE ROWNUM <={numb}"
        
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()
        collums = getCollumNames()
        a=collums.description
     
        for i in range(0, numb):
        # Create a dictionary to store the data for each i
            i_data = {}
            for j in range(len(a)):
                key = a[j][0]
                value = records[i][j]

                if key in i_data:
                    i_data[key].append(value)  # If the key already exists, append the new value
                else:
                    i_data[key] = value  # If the key doesn't exist, create a list with the value

            # Add the i_data dictionary to the consulta dictionary under the i key
            consulta[i] = i_data

        
        #print(consulta)

        # print(f"\n <oracledb.LOB object at 0x7f8823d022b0> \n {consulta['OBJETIVOS']} \n")
        #consulta[0]['OBJETIVOS'] = str(consulta[0]['OBJETIVOS'])
            
        cursor.close()

    except oracledb.Error as error:
        print("Failed to read data from table", error)
    finally:
        if connection:
            connection.close()
            print("The connection is closed")
    
    # return records
    return consulta


# a = int(input(""))
# lul = getlimitedRows(a)
# print("\n")
# print(type(lul))

def getallRows():
   
    try:
        connection = oracledb.connect(conStr)
        cursor = connection.cursor()
        print("Connected to database")
        sqlite_select_query = f"SELECT * FROM IDEA.STG_PROJETOS_CONVENIAR"
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()
        length = len(records)
        print(len(records))
        cursor.execute(sqlite_select_query)
       
        cursor.close()

    except oracledb.Error as error:
        print("Failed to read data from table", error)
    finally:
        if connection:
            connection.close()
            print("The connection is closed")
    
    # return records
    return length



