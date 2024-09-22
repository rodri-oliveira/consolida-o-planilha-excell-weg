import os
import pandas as pd
from tkinter import messagebox

# Variável global para armazenar o DataFrame consolidado
dataframe = None

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_pasta, caminho_saida):
    global dataframe
    lista_dataframes = []

    # Itera sobre todos os arquivos da pasta especificada
    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo.endswith('.xlsx'):  # Verifica se o arquivo é um Excel
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)
            xls = pd.ExcelFile(caminho_arquivo_excel)  # Carrega o arquivo Excel

            # Itera sobre todas as abas (sheets) do Excel
            for sheet_name in xls.sheet_names:
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)
                lista_dataframes.append(dataframe_planilha)

    # Se houver planilhas carregadas, concatena e salva
    if lista_dataframes:
        dataframe = pd.concat(lista_dataframes, ignore_index=True)
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")

# Função para gerar um relatório filtrado com base nos critérios
def gerar_relatorio(dataframe, coluna, operador, valor):
    try:
        if coluna not in dataframe.columns:
            messagebox.showwarning("Erro", f"A coluna '{coluna}' não existe no DataFrame.")
            return pd.DataFrame()  # Retorna um DataFrame vazio se a coluna não for encontrada

        # Converte a coluna para numérico, ignorando erros de conversão
        dataframe[coluna] = pd.to_numeric(dataframe[coluna], errors='coerce')

        # Aplica o filtro com base no operador
        if operador == "maior que":
            return dataframe[dataframe[coluna] > valor]
        elif operador == "menor que":
            return dataframe[dataframe[coluna] < valor]
        elif operador == "igual a":
            return dataframe[dataframe[coluna] == valor]
        return pd.DataFrame()  # Retorna um DataFrame vazio se não houver correspondência
    except Exception as e:
        messagebox.showwarning("Erro", f"Ocorreu um erro ao aplicar o filtro: {str(e)}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro
