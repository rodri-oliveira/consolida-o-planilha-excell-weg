import pandas as pd
import os

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_das_planilhas):
    """
    Consolida todas as planilhas de todos os arquivos Excel no diretório especificado,
    criando colunas específicas como MÊS, ANO, Epic, Status, Due Date, Assignee e horas.

    Parâmetros:
    caminho_das_planilhas (str): Caminho do diretório onde estão as planilhas a serem consolidadas.
    """

    global dataframe_consolidado
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):  # Verifica se o arquivo tem a extensão correta
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            
            # Carrega o arquivo Excel
            xls = pd.ExcelFile(caminho_completo)

            # Itera sobre todas as abas do arquivo Excel
            for nome_aba in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verifica se o DataFrame contém as colunas necessárias
                if 'Planned effort' in df.columns:
                    # Criação das colunas 'MÊS' e 'ANO'
                    df['MÊS'] = df.columns[8]  # Coluna 'I' corresponde ao índice 8
                    df['ANO'] = df.columns[0]  # Coluna 'A' corresponde ao índice 0

                    # Seleciona as colunas necessárias para o formato final
                    colunas_necessarias = ['Epic', 'Status', 'Due Date', 'Assignee', 'Planned effort', 'MÊS', 'ANO']
                    
                    if all(col in df.columns for col in colunas_necessarias):
                        # Renomeia a coluna 'Planned effort' para 'horas'
                        df = df[colunas_necessarias].rename(columns={'Planned effort': 'horas'})

                        # Adiciona o DataFrame à lista de DataFrames a serem consolidados
                        lista_dfs.append(df)
                    else:
                        print(f"Aba {nome_aba} do arquivo {arquivo} não contém todas as colunas esperadas.")
                else:
                    print(f"Aba {nome_aba} do arquivo {arquivo} não contém a coluna 'Planned effort'.")

    # Concatena todos os DataFrames da lista em um único DataFrame consolidado
    if lista_dfs:
        dataframe_consolidado = pd.concat(lista_dfs, ignore_index=True)

        # Salvando o DataFrame consolidado no caminho especificado
        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")
