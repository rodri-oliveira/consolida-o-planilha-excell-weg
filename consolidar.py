import os  # Certifique-se de adicionar isso para trabalhar com diretórios e arquivos
import pandas as pd

import os  # Certifique-se de adicionar isso para trabalhar com diretórios e arquivos
import pandas as pd


def consolidar_planilhas(caminho_das_planilhas):
    """Consolida todas as planilhas de todos os arquivos Excel no diretório especificado, ignorando a aba 'Backlog'."""
    global dataframe_consolidado
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            print(f"Processando arquivo: {arquivo}")
            for nome_aba in xls.sheet_names:
                # Verifica se a aba é "Backlog" e ignora
                if 'Backlog' in nome_aba:
                    print(f"Ignorando aba: '{nome_aba}'")
                    continue

                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verifica se a aba tem as colunas mínimas necessárias
                if 'Planned effort' in df.columns and df.shape[1] > 5:
                    # Obtendo o valor da "Seção" que está duas linhas acima da célula da "Equipe"
                    valor_secao = df.iloc[3, 0]  # Captura a célula A5 (duas linhas acima da célula de equipe)
                    valor_equipe = df.iloc[4, 0]  # Captura a célula A6 (abaixo da seção)

                    for index, row in df.iterrows():
                        if index < 3:  # Ignora as primeiras linhas
                            continue

                        # Captura os dados relevantes
                        epic = row['Epic'] if 'Epic' in df.columns else ''
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        planned_effort = row['Planned effort']
                        estimate_effort = row.iloc[5] if len(row) > 5 else None

                        # As colunas de meses são a partir da coluna I (coluna index 8)
                        colunas_meses = df.columns[8:]  # Assume que as primeiras 8 colunas são fixas

                        for idx, mes in enumerate(colunas_meses):
                            try:
                                valor_hora_mes = row[mes]

                                # Verifica se o valor da célula é válido e se é numérico
                                if pd.isnull(valor_hora_mes) or not isinstance(valor_hora_mes, (int, float)):
                                    continue

                                # Verifica se a coluna do mês contém um valor esperado
                                if isinstance(mes, str) and '/' in mes:
                                    mes_abreviado, ano_abreviado = mes.split('/')
                                else:
                                    print(f"Erro no arquivo '{arquivo}', aba '{nome_aba}', linha {index + 1}, coluna '{mes}': formato inesperado para o mês.")
                                    continue

                                # Lógica para determinar o ano completo com base na posição do mês
                                ano = int('20' + ano_abreviado)

                                nova_linha = {
                                    'Epic': epic,
                                    'Status': status,
                                    'Due Date': due_date,
                                    'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                                    'Planned Effort': planned_effort,
                                    'Estimate Effort': estimate_effort,
                                    'MÊS': mes_abreviado,
                                    'ANO': ano,
                                    'Horas mês': valor_hora_mes,
                                    'Seção': valor_secao,    # Adiciona o valor da Seção
                                    'Equipe': valor_equipe   # Adiciona o valor da Equipe
                                }
                                lista_dfs.append(nova_linha)

                            except Exception as e:
                                print(f"Erro no arquivo '{arquivo}', aba '{nome_aba}', linha {index + 1}, coluna '{mes}': {str(e)}")

    # Consolida os dados em um DataFrame
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Exclui colunas indesejadas
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        # Define o caminho para salvar a planilha consolidada
        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")

def consolidar_aba_backlog(caminho_das_planilhas):
    """Consolida as planilhas da aba 'Backlog' de todos os arquivos Excel no diretório especificado."""
    global dataframe_consolidado
    lista_dfs = []

    # Itera sobre todos os arquivos no diretório especificado
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            print(f"Processando arquivo: {arquivo}")

            # Processa cada aba do arquivo Excel
            for nome_aba in xls.sheet_names:
                # Verifica se a aba é "Backlog"
                if 'Backlog' in nome_aba:
                    print(f"Processando aba: '{nome_aba}'")

                    # Lê a aba "Backlog" como DataFrame
                    df = pd.read_excel(xls, sheet_name=nome_aba)

                    # Verifica se as colunas necessárias existem e o DataFrame tem colunas suficientes
                    if 'Estimated effort' in df.columns and df.shape[1] > 5:
                        # Obtém os valores da "Seção" e "Equipe"
                        valor_secao = df.iloc[3, 0]  # Captura a célula A4
                        valor_equipe = df.iloc[4, 0]  # Captura a célula A5

                        # Itera sobre as linhas do DataFrame
                        for index, row in df.iterrows():
                            if index < 5:  # Ignora as primeiras linhas
                                continue

                            # Captura os dados relevantes
                            epic = row['Epic'] if 'Epic' in df.columns else ''
                            status = row['Status'] if 'Status' in df.columns else ''
                            due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                            estimated_effort = row['Estimated effort']  # Valor da coluna Estimated effort

                            # Cria a nova linha para o DataFrame consolidado
                            nova_linha = {
                                'Epic': epic,
                                'Status': status,
                                'Due Date': due_date,
                                'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                                'Estimated Effort': estimated_effort,
                                'Seção': valor_secao,    # Adiciona o valor da Seção
                                'Equipe': valor_equipe   # Adiciona o valor da Equipe
                            }
                            lista_dfs.append(nova_linha)

    # Consolida os dados em um DataFrame
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Exclui colunas indesejadas
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        # Exclui linhas onde a primeira coluna está em branco
        dataframe_consolidado = dataframe_consolidado[dataframe_consolidado.iloc[:, 0].notna()]

        # Define o caminho para salvar a planilha consolidada
        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/backlog-consolidado/backlog_consolidado.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")


def consolidar_horas_backlog(caminho_das_planilhas):
    """Consolida os dados das abas 'Backlog' de todos os arquivos Excel no diretório especificado."""
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            print(f"Processando arquivo: {arquivo}")
            for nome_aba in xls.sheet_names:
                # Verifica se a aba é "Backlog" e processa apenas ela
                if 'Backlog' in nome_aba:
                    print(f"Processando aba: '{nome_aba}'")
                    df = pd.read_excel(xls, sheet_name=nome_aba)

                    # Itera pelas linhas de "Epic" e gera os dados de "Hora/mês"
                    for index, row in df.iterrows():
                        # Pega o valor do Epic (coluna A), Planned effort (coluna G), e Hora/mês (colunas H a S)
                        epic = row[0]  # Valor da coluna A (Epic)
                        planned_effort = row[6]  # Valor da coluna G (Planned effort)

                        # Verifica se há dados nas colunas de H a S (12 meses)
                        meses = ['Mês 1', 'Mês 2', 'Mês 3', 'Mês 4', 'Mês 5', 'Mês 6', 'Mês 7', 'Mês 8', 'Mês 9', 'Mês 10', 'Mês 11', 'Mês 12']
                        horas_mes = row[7:19]  # Colunas H até S

                        # Verifica se as células contêm dados válidos
                        if not horas_mes.isnull().all():
                            # Cria uma linha para cada mês, com o respectivo valor de Hora/mês
                            for i, hora in enumerate(horas_mes):
                                nova_linha = {
                                    'Epic': epic,
                                    'Planned effort': planned_effort,
                                    'Hora/mês': hora,
                                    'Mês': meses[i]
                                }
                                lista_dfs.append(nova_linha)

    # Consolida os dados em um DataFrame final
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Define o caminho para salvar a planilha consolidada
        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/horas-backlog/consolidada-backlog-horas.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma aba 'Backlog' foi encontrada para consolidar.")
