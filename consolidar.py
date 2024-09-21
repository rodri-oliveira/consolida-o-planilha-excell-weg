import os
import pandas as pd

def consolidar_planilhas(caminho_pasta):
    # Mapeamento das planilhas e suas respectivas abas
    abas = {
        'controle-estoque1.xlsx': 'controle-estoque-1',
        'controle-estoque2.xlsx': 'controle-estoque-2',
        'controle-estoque3.xlsx': 'controle-estoque-3'
    }

    lista_dataframes = []
    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo in abas:
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)
            nome_aba = abas[nome_arquivo]
            dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba)
            lista_dataframes.append(dataframe_planilha)

    if lista_dataframes:
        dataframe_consolidado = pd.concat(lista_dataframes, ignore_index=True)
        caminho_saida_excel = os.path.join(caminho_pasta, 'planilha_consolidada.xlsx')
        dataframe_consolidado.to_excel(caminho_saida_excel, index=False)
    else:
        raise ValueError("Nenhuma planilha foi consolidada.")
