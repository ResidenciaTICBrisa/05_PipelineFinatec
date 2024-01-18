import oracledb

#connection string in the format
#<username>/<password>@<dBhostAddress>:<dbPort>/<dbServiceName>

conStr = ''

def getNomeColunas():

    #inicializando o objeto que ira conectar no db
    conn = None
    #criando o objeto de conexão das
    conn = oracledb.connect(conStr)
    #criar um objeto cursor necessario para fazer as consultas
    cur = conn.cursor() 
    cur.execute("SELECT * FROM IDEA.STG_PROJETOS_CONVENIAR")

    return cur

def consultaPorID(IDPROJETO):
    try:
        connection = oracledb.connect(conStr)
        cursor = connection.cursor()
        print("Connected to database")

        # idProjeto = 6411
        sqlite_select_query = f"SELECT * FROM IDEA.STG_PROJETOS_CONVENIAR WHERE CODIGO='{IDPROJETO}'"
        
        cursor.execute(sqlite_select_query)

        records = cursor.fetchall()

        collums = getNomeColunas()
        # print(records)

        # print(len(collums.description))
        # print(len(records[0]))

        consulta = {}

        for i in range(len(collums.description)):
            # print(f"{collums.description[i][0]}: {records[0][i]}")
            consulta[collums.description[i][0]] = records[0][i]

        print(consulta)

        # print(f"\n <oracledb.LOB object at 0x7f8823d022b0> \n {consulta['OBJETIVOS']} \n")
        consulta['OBJETIVOS'] = str(consulta['OBJETIVOS'])
            
        cursor.close()

    except oracledb.Error as error:
        print("Failed to read data from table", error)
    finally:
        if connection:
            connection.close()
            print("The connection is closed")
    
    # return records
    return consulta

# executando
a = input("Digite o id do projeto: ")
print("\n\n")
resultado = consultaPorID(a)

print(f"\n {resultado} \n")



