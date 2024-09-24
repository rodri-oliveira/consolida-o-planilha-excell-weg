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
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)
            xls = pd.ExcelFile(caminho_arquivo_excel)

            # Adiciona cada aba como um DataFrame na lista
            for sheet_name in xls.sheet_names:
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)
                lista_dataframes.append(dataframe_planilha)

    # Concatena todos os DataFrames em um único DataFrame
    if lista_dataframes:
        dataframe = pd.concat(lista_dataframes, ignore_index=True)

        # Salva o DataFrame consolidado em um arquivo Excel
        try:
            os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)  # Cria o diretório de saída se não existir
            dataframe.to_excel(caminho_saida, index=False)
            print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar a planilha: {str(e)}")
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")
