import os
import pandas as pd

def consolidar_planilhas():
    # Caminho para a pasta onde as planilhas do Excel estão localizadas
    caminho_pasta_planilhas = r'C:\consolidar-planilha-weg\planilas-mari-weg'
    
    # Lista para armazenar os DataFrames de cada planilha
    lista_dataframes = []
    
    # Iterar sobre os arquivos na pasta de planilhas
    for nome_arquivo in os.listdir(caminho_pasta_planilhas):
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo_excel = os.path.join(caminho_pasta_planilhas, nome_arquivo)
            
            # Nome da aba será o mesmo nome do arquivo, sem a extensão '.xlsx'
            nome_aba = nome_arquivo.replace('.xlsx', '')
            
            # Ler a planilha Excel para um DataFrame, usando o nome do arquivo como nome da aba
            dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba)
            
            # Adicionar o DataFrame à lista
            lista_dataframes.append(dataframe_planilha)
    
    # Combinar todos os DataFrames em um único DataFrame
    dataframe_consolidado = pd.concat(lista_dataframes, ignore_index=True)
    
    # Caminho para salvar o arquivo consolidado
    caminho_saida_excel = os.path.join(caminho_pasta_planilhas, 'planilha_consolidada.xlsx')
    
    # Salvar o DataFrame consolidado em um novo arquivo Excel
    dataframe_consolidado.to_excel(caminho_saida_excel, index=False)
    
    # Mensagem de sucesso
    print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida_excel}')

if __name__ == "__main__":
    consolidar_planilhas()

