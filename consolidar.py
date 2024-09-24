import os  # Importa o módulo para interagir com o sistema operacional
import pandas as pd  # Importa a biblioteca pandas para manipulação de dados

def consolidar_planilhas(caminho_pasta_origem):
    """
    Consolida todas as planilhas em arquivos Excel (.xlsx) de uma pasta em um único DataFrame.

    Args:
        caminho_pasta_origem (str): Caminho da pasta contendo os arquivos Excel.

    Returns:
        pd.DataFrame: DataFrame consolidado com os dados de todas as planilhas.
    """
    lista_dataframes = []  # Lista para armazenar os DataFrames das planilhas

    try:
        # Itera sobre todos os arquivos na pasta de origem
        for nome_arquivo in os.listdir(caminho_pasta_origem):
            # Verifica se o arquivo é um arquivo Excel
            if nome_arquivo.endswith('.xlsx'):
                # Cria o caminho completo do arquivo Excel
                caminho_arquivo_excel = os.path.join(caminho_pasta_origem, nome_arquivo)
                xls = pd.ExcelFile(caminho_arquivo_excel)  # Lê o arquivo Excel

                # Itera sobre todas as abas do arquivo
                for nome_aba in xls.sheet_names:
                    # Lê a aba e armazena o DataFrame na lista
                    dataframe_aba = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba)
                    lista_dataframes.append(dataframe_aba)

        # Verifica se há DataFrames para consolidar
        if lista_dataframes:
            # Concatena todos os DataFrames em um único DataFrame
            dataframe_consolidado = pd.concat(lista_dataframes, ignore_index=True)
            return dataframe_consolidado  # Retorna o DataFrame consolidado
        else:
            print("Nenhuma planilha foi encontrada ou consolidada.")
            return pd.DataFrame()  # Retorna um DataFrame vazio

    except FileNotFoundError:
        print(f"O caminho '{caminho_pasta_origem}' não foi encontrado.")
        return pd.DataFrame()  # Retorna um DataFrame vazio
    except pd.errors.EmptyDataError:
        print("Um dos arquivos Excel está vazio.")
        return pd.DataFrame()  # Retorna um DataFrame vazio
    except Exception as e:
        print(f"Ocorreu um erro ao consolidar as planilhas: {str(e)}")
        return pd.DataFrame()  # Retorna um DataFrame vazio

def salvar_planilha(dataframe, caminho_saida):
    """
    Salva um DataFrame em um arquivo Excel.

    Args:
        dataframe (pd.DataFrame): DataFrame a ser salvo.
        caminho_saida (str): Caminho do arquivo de saída onde o DataFrame será salvo.

    Returns:
        None
    """
    try:
        # Cria a pasta de saída se não existir
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
        # Salva o DataFrame em um arquivo Excel
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilha salva em: {caminho_saida}')
    except FileNotFoundError:
        print("O caminho para salvar a planilha não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro ao salvar a planilha: {str(e)}")
