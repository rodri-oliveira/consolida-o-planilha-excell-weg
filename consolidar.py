import pandas as pd
import os

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_das_planilhas):
    """
    Consolida todas as planilhas de todos os arquivos Excel no diretório especificado,
    criando colunas específicas como MÊS, ANO, Epic, Status, Due Date, Assignee e Planned Effort.
    
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
                # Verifica se a aba é "Backlog"
                if nome_aba == "Backlog":
                    print(f"Aba '{nome_aba}' do arquivo {arquivo} foi ignorada.")
                    continue  # Pula para a próxima aba

                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verifica se o DataFrame contém as colunas necessárias
                if 'Planned effort' in df.columns:
                    # Loop para preencher os meses e anos
                    for index, row in df.iterrows():
                        epic = row['Epic'] if 'Epic' in df.columns else ''
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        assignee = row['Assignee'] if 'Assignee' in df.columns else ''
                        planned_effort = row['Planned effort']

                        # Verifique se as colunas de I a Y estão presentes no DataFrame
                        colunas_meses = df.columns[8:25]  # I até Y (colunas 9 até 25)

                        # Adiciona uma linha para cada mês de agosto de 2024 até dezembro de 2025
                        for idx, mes in enumerate(colunas_meses):
                            valor_hora_mes = row[mes]  # Pegando o valor da célula para o mês correspondente

                            ano = 2024 if idx < 5 else 2025  # Determina o ano: 2024 para os primeiros 5 meses, depois 2025

                            nova_linha = {
                                'Epic': epic,
                                'Status': status,
                                'Due Date': due_date,
                                'Assignee': assignee,
                                'Planned Effort': planned_effort,  # A coluna "Planned Effort"
                                'MÊS': mes,
                                'ANO': ano,
                                'Horas mês': valor_hora_mes  # Valor das horas do mês correspondente (I5 até Y5)
                            }
                            lista_dfs.append(nova_linha)

                else:
                    print(f"Aba {nome_aba} do arquivo {arquivo} não contém a coluna 'Planned effort'.")

    # Cria um DataFrame a partir da lista de dicionários
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Renomear a coluna "horas" para "Planned Effort"
        dataframe_consolidado.rename(columns={'horas': 'Planned Effort'}, inplace=True)

        # Salvando o DataFrame consolidado no caminho especificado
        caminho_para_salvar_arquivo = 'C:/consolida-o-planilha-excell-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")
