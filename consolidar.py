import os  # Módulo para interagir com o sistema de arquivos
import pandas as pd  # Biblioteca para manipulação de dados em DataFrames

def consolidar_planilhas(caminho_pasta, caminho_saida):
    global dataframe  # Declara a variável dataframe como global para acessá-la fora da função
    lista_dataframes = []  # Cria uma lista para armazenar os DataFrames de cada planilha

    # Itera sobre todos os arquivos na pasta especificada
    for nome_arquivo in os.listdir(caminho_pasta):
        # Verifica se o arquivo é uma planilha Excel
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)  # Cria o caminho completo do arquivo
            xls = pd.ExcelFile(caminho_arquivo_excel)  # Carrega o arquivo Excel

            # Itera sobre as abas (sheets) do arquivo Excel
            for sheet_name in xls.sheet_names:
                # Lê a aba específica em um DataFrame
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)
                lista_dataframes.append(dataframe_planilha)  # Adiciona o DataFrame à lista

    # Verifica se a lista de DataFrames não está vazia
    if lista_dataframes:
        # Combina todos os DataFrames em um único DataFrame
        dataframe = pd.concat(lista_dataframes, ignore_index=True)
        # Salva o DataFrame consolidado em um novo arquivo Excel
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')  # Mensagem de sucesso
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")  # Mensagem de erro se não houver planilhas
