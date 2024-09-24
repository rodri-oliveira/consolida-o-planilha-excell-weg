import os
import pandas as pd

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_pasta):
    global dataframe
    lista_dataframes = []

    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)
            xls = pd.ExcelFile(caminho_arquivo_excel)

            for sheet_name in xls.sheet_names:
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)
                lista_dataframes.append(dataframe_planilha)

    if lista_dataframes:
        dataframe = pd.concat(lista_dataframes, ignore_index=True)
        caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx"
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")
