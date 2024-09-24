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

<<<<<<< HEAD
        # Salva o DataFrame consolidado em um arquivo Excel
        try:
            os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)  # Cria o diretório de saída se não existir
            dataframe.to_excel(caminho_saida, index=False)
            print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar a planilha: {str(e)}")
=======
        # Verifica e imprime as colunas do DataFrame consolidado
        print("Colunas do DataFrame consolidado:", dataframe.columns.tolist())

        # Define o caminho de saída para as planilhas consolidadas
        caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx"
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
>>>>>>> 6fb6f02ffa78154ca29c6a10547c683986b3db03
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")
