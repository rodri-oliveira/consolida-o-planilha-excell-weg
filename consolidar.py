import pandas as pd
import os

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_das_planilhas, caminho_para_salvar):
    """
    Consolida todas as planilhas do diretório especificado, organizando-as em um único arquivo de saída.

    Parâmetros:
    caminho_das_planilhas (str): Caminho do diretório onde estão as planilhas a serem consolidadas.
    caminho_para_salvar (str): Caminho do diretório onde o arquivo consolidado será salvo.
    """

    # Lista para armazenar os dados de cada planilha
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):  # Verifica se o arquivo tem a extensão correta
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            
            # Lê a planilha Excel
            df = pd.read_excel(caminho_completo)

            # Reorganizando as colunas conforme o formato desejado
            colunas_necessarias = ['Epic', 'Ações', 'Status', 'Due Date', 'Assignee',
                                   'Estimated effort', 'Planned effort', 'Ano', 'Mês', 'horas']
            
            # Verifica se todas as colunas necessárias estão no DataFrame
            if all(col in df.columns for col in colunas_necessarias):
                # Seleciona apenas as colunas necessárias
                df = df[colunas_necessarias]
                lista_dfs.append(df)
            else:
                print(f"Arquivo {arquivo} não contém as colunas esperadas. Ignorando esse arquivo.")

    # Concatena todos os DataFrames da lista em um único DataFrame consolidado
    if lista_dfs:
        df_consolidado = pd.concat(lista_dfs, ignore_index=True)

        # Salvando o DataFrame consolidado no caminho especificado
        caminho_para_salvar_arquivo = os.path.join(caminho_para_salvar, 'relatorio_consolidado.xlsx')
        df_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")
