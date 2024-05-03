import os

def deletar_arquivos_em_pasta(caminho_da_pasta):
    # Verificar se o caminho é um diretório
    if os.path.isdir(caminho_da_pasta):
        # Listar todos os arquivos na pasta
        arquivos = os.listdir(caminho_da_pasta)
        # Imprimir o nome de cada arquivo
        for arquivo in arquivos:
            # Verificar se o arquivo é um arquivo XLSX
            if arquivo.endswith('.xlsx'):
                # Construir o caminho completo do arquivo
                caminho_arquivo = os.path.join(caminho_da_pasta, arquivo)
                # Tentar excluir o arquivo
                try:
                    os.remove(caminho_arquivo)
                    print(f"Arquivo {arquivo} excluído com sucesso.")
                except OSError as e:
                    print(f"Erro ao excluir o arquivo {arquivo}: {e}")
            
    else:
        print("O caminho especificado não é um diretório.")

# # Caminho da pasta de planilhas preenchidas
# caminhoPastaPlanilhasPreenchidas = "../../planilhas_preenchidas/"


# #
# diretorio_atual = os.path.dirname(os.path.abspath(__file__))
# testeCaminho = os.path.join(diretorio_atual, caminhoPastaPlanilhasPreenchidas)

# # Chamar a função para listar e imprimir os arquivos na pasta
# listar_arquivos_em_pasta(testeCaminho)