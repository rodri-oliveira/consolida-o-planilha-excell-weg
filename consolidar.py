import os
import pandas as pd
import unidecode
from io import BytesIO
from auth import enviar_para_sharepoint, obter_token_sharepoint, buscar_arquivos_pasta
from dotenv import load_dotenv
import requests
import pandas as pd
from io import BytesIO
import csv

load_dotenv()
token = obter_token_sharepoint()
SITE_URL = os.getenv('SITE_URL')

def limpar_string(s):
    """Remove caracteres especiais de uma string, permitindo apenas alfanuméricos, espaços, sublinhados e hifens."""
    # Verifica se 's' é uma string
    if isinstance(s, str):
        return ''.join(e for e in s if e.isalnum() or e in [' ', '_', '-'])
    return ''  # Retorna uma string vazia se 's' não for uma string


# Função principal para consolidar planilhas do SharePoint

def consolidar_planilhas_sharepoint(lista_arquivos_sharepoint, token):
    """Consolida todas as planilhas dos arquivos Excel do SharePoint, ignorando a aba 'Backlog'."""
    
    lista_dfs = []
    
    # Loop para percorrer todos os arquivos no SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")
        
        # Montar endpoint do arquivo
        END_POINT = f"https://weg365.sharepoint.com/teams/BR-TI-TIN/_api/Web/GetFileByServerRelativePath(decodedurl='{url_arquivo}')/$value"
        response = requests.get(END_POINT, headers={'Authorization': f'Bearer {token}'})
        
        if response.status_code != 200:
            print(f"Erro ao acessar arquivo: {url_arquivo} - {response.status_code}")
            continue
        
        # Lê o conteúdo da resposta no formato Excel
        arquivo_excel = BytesIO(response.content)
        xls = pd.ExcelFile(arquivo_excel)
        
        for nome_aba in xls.sheet_names:
            # Ignora a aba "Backlog"
            if 'Backlog' in nome_aba:
                print(f"Ignorando aba: '{nome_aba}'")
                continue

            df = pd.read_excel(xls, sheet_name=nome_aba)

            # Renomeia colunas (remove acentos e substitui espaços por underscores)
            df.columns = [unidecode.unidecode(col).replace(" ", "_") for col in df.columns]

            # Verifica se a aba tem as colunas mínimas necessárias
            if 'Planned_effort' in df.columns and df.shape[1] > 5:
                valor_secao = df.iloc[3, 0]  # Captura a célula A4
                valor_equipe = df.iloc[4, 0]  # Captura a célula A5
                
                # Captura os valores de G1, G2 e G3
                valor_gap = df.iloc[0, 6]  # Célula G1
                valor_horas_disponiveis = df.iloc[1, 6]  # Célula G2
                valor_total_esforco = df.iloc[2, 6]  # Célula G3

                for index, row in df.iterrows():
                    if index < 3:  # Ignora as primeiras linhas
                        continue

                    epic = limpar_string(row['Epic']) if 'Epic' in df.columns else ''
                    status = limpar_string(row['Status']) if 'Status' in df.columns else ''
                    due_date = limpar_string(row['Due_Date']) if 'Due_Date' in df.columns else ''
                    planned_effort = row['Planned_effort']
                    estimate_effort = row.iloc[5] if len(row) > 5 else None

                    colunas_meses = df.columns[8:]

                    for idx, mes in enumerate(colunas_meses):
                        try:
                            valor_hora_mes = row[mes]

                            if pd.isnull(valor_hora_mes) or not isinstance(valor_hora_mes, (int, float)):
                                continue

                            if isinstance(mes, str) and '/' in mes:
                                mes_abreviado, ano_abreviado = mes.split('/')
                                ano = int('20' + ano_abreviado)
                            else:
                                print(f"Erro no arquivo '{url_arquivo}', aba '{nome_aba}', linha {index + 1}, coluna '{mes}': formato inesperado para o mês.")
                                continue

                            nova_linha = {
                                'Epic': epic,
                                'Status': status,
                                'Due_Date': due_date,
                                'Assignee': limpar_string(row['Assignee']) if 'Assignee' in df.columns else '',
                                'Planned_Effort': planned_effort,
                                'Estimate_Effort': estimate_effort,
                                'MES': mes_abreviado,
                                'ANO': ano,
                                'Horas_mes': valor_hora_mes,
                                'Secao': valor_secao,
                                'Equipe': valor_equipe,
                                'GAP': valor_gap,  # Adiciona a coluna GAP
                                'Horas_disponiveis': valor_horas_disponiveis,  # Adiciona a coluna Horas_disponiveis
                                'Total_esforco': valor_total_esforco  # Adiciona a coluna Total_esforco
                            }
                            lista_dfs.append(nova_linha)

                        except Exception as e:
                            print(f"Erro no arquivo '{url_arquivo}', aba '{nome_aba}', linha {index + 1}, coluna '{mes}': {str(e)}")

    # Consolida os dados em um DataFrame
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_csv(arquivo_memoria, index=False, quoting=csv.QUOTE_NONE)  # Salva como CSV sem aspas
        arquivo_memoria.seek(0)

        # Nome do arquivo a ser salvo no SharePoint
        nome_arquivo_sharepoint = "planilha_consolidada.csv"

        try:
            # Faz o upload do arquivo diretamente da memória para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint)
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")



# Função principal para consolidar a aba "Backlog" e enviar para o SharePoint

def consolidar_aba_backlog_sharepoint(lista_arquivos_sharepoint, token):
    """Consolida as planilhas da aba 'Backlog' de todos os arquivos Excel no SharePoint."""

    global dataframe_consolidado
    lista_dfs = []

    # Itera sobre todos os arquivos na lista de arquivos do SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")

        # Montar endpoint do arquivo
        END_POINT = f"https://weg365.sharepoint.com/teams/BR-TI-TIN/_api/Web/GetFileByServerRelativePath(decodedurl='{url_arquivo}')/$value"

        # Usa requests para acessar diretamente o arquivo no SharePoint
        response = requests.get(END_POINT, headers={'Authorization': f'Bearer {token}'})

        if response.status_code != 200:
            print(f"Erro ao acessar arquivo: {url_arquivo} - {response.status_code}")
            continue

        # Usa o conteúdo diretamente, sem baixar
        arquivo_excel = BytesIO(response.content)
        xls = pd.ExcelFile(arquivo_excel)

        # Processa cada aba do arquivo Excel
        for nome_aba in xls.sheet_names:
            # Verifica se a aba é "Backlog"
            if 'Backlog' in nome_aba:
                print(f"Processando aba: '{nome_aba}'")

                # Lê a aba "Backlog" como DataFrame
                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Renomeia colunas (remove acentos e substitui espaços por underscores)
                df.columns = [unidecode.unidecode(col).replace(" ", "_") for col in df.columns]

                # Verifica se as colunas necessárias existem e o DataFrame tem colunas suficientes
                if 'Estimated_effort' in df.columns and df.shape[1] > 5:
                    # Obtém os valores da "Seção" e "Equipe"
                    valor_secao = df.iloc[3, 0]  # Captura a célula A4
                    valor_equipe = df.iloc[4, 0]  # Captura a célula A5

                    # Itera sobre as linhas do DataFrame
                    for index, row in df.iterrows():
                        if index < 5:  # Ignora as primeiras linhas
                            continue

                        # Captura os dados relevantes
                        epic = limpar_string(row['Epic']) if 'Epic' in df.columns else ''
                        status = limpar_string(row['Status']) if 'Status' in df.columns else ''
                        due_date = limpar_string(row['Due_Date']) if 'Due_Date' in df.columns else ''
                        estimated_effort = row['Estimated_effort']  # Valor da coluna Estimated effort

                        # Cria a nova linha para o DataFrame consolidado
                        nova_linha = {
                            'Epic': epic,
                            'Status': status,
                            'Due_Date': due_date,
                            'Assignee': limpar_string(row['Assignee']) if 'Assignee' in df.columns else '',
                            'Estimated_Effort': estimated_effort,
                            'Secao': valor_secao,    # Adiciona o valor da Seção
                            'Equipe': valor_equipe   # Adiciona o valor da Equipe
                        }
                        lista_dfs.append(nova_linha)

    # Consolida os dados em um DataFrame
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Exclui colunas indesejadas
        dataframe_consolidado.drop(columns=['Horas_disponíveis', 'Total_de_esforço'], inplace=True, errors='ignore')

        # Exclui linhas onde a primeira coluna está em branco
        dataframe_consolidado = dataframe_consolidado[dataframe_consolidado.iloc[:, 0].notna()]

        # Define o caminho para salvar a planilha consolidada
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        nome_arquivo_sharepoint = "backlog_consolidado.csv"  # Mudança aqui para .csv

        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_csv(arquivo_memoria, index=False, quoting=csv.QUOTE_NONE)  # Salva como CSV sem aspas
        arquivo_memoria.seek(0)  # Move o ponteiro de volta ao início do arquivo

        try:
            # Faz o upload do arquivo diretamente da memória para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint)
            print(f"Relatório consolidado salvo em: {url_diretorio_sharepoint}/{nome_arquivo_sharepoint}")
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")

    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")

def consolidar_horas_backlog_sharepoint(lista_arquivos_sharepoint, token):
    """Consolida os dados das abas 'Backlog' de todos os arquivos Excel no SharePoint."""
    lista_dfs = []

    # Loop para percorrer todos os arquivos na lista de arquivos do SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")

        # Montar endpoint do arquivo
        END_POINT = f"https://weg365.sharepoint.com/teams/BR-TI-TIN/_api/Web/GetFileByServerRelativePath(decodedurl='{url_arquivo}')/$value"

        # Faz a requisição para acessar o arquivo no SharePoint
        response = requests.get(END_POINT, headers={'Authorization': f'Bearer {token}'})

        if response.status_code != 200:
            print(f"Erro ao acessar arquivo: {url_arquivo} - {response.status_code}")
            continue

        # Usa o conteúdo diretamente, sem baixar
        arquivo_excel = BytesIO(response.content)
        xls = pd.ExcelFile(arquivo_excel)

        # Processa cada aba do arquivo Excel
        for nome_aba in xls.sheet_names:
            # Verifica se a aba contém "Backlog" e processa apenas ela
            if 'Backlog' in nome_aba:
                print(f"Processando aba: '{nome_aba}'")
                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Renomeia colunas (remove acentos e substitui espaços por underscores)
                df.columns = [unidecode.unidecode(col).replace(" ", "_") for col in df.columns]

                # Itera pelas linhas de "Epic" e gera os dados de "Hora_mes"
                for index, row in df.iterrows():
                    # Pega o valor do Epic (coluna A), Planned effort (coluna G), e Hora_mes (colunas H a S)
                    epic = limpar_string(row.iloc[0])  # Valor da coluna A (Epic) usando iloc
                    planned_effort = row.iloc[6]  # Valor da coluna G (Planned effort)

                    # Verifica se há dados nas colunas de H a S (12 meses)
                    meses = ['Mes_1', 'Mes_2', 'Mes_3', 'Mes_4', 'Mes_5', 'Mes_6', 
                             'Mes_7', 'Mes_8', 'Mes_9', 'Mes_10', 'Mes_11', 'Mes_12']
                    horas_mes = row.iloc[7:19]  # Colunas H até S usando iloc

                    # Verifica se as células contêm dados válidos
                    if not horas_mes.isnull().all():
                        # Cria uma linha para cada mês, com o respectivo valor de Hora_mes
                        for i, hora in enumerate(horas_mes):
                            nova_linha = {
                                'Epic': epic,
                                'Planned_effort': planned_effort,
                                'Hora_mes': hora,
                                'Mes': meses[i]
                            }
                            lista_dfs.append(nova_linha)

    # Consolida os dados em um DataFrame final
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Define o caminho para salvar a planilha consolidada
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        nome_arquivo_sharepoint = "consolidada-backlog-horas.csv"
        
        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_csv(arquivo_memoria, index=False)  # Salvando como CSV
        arquivo_memoria.seek(0)  # Move o ponteiro de volta ao início do arquivo

        try:
            # Faz o upload do arquivo diretamente da memória para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint)
            print(f"Relatório consolidado salvo em: {url_diretorio_sharepoint}/{nome_arquivo_sharepoint}")
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")
    else:
        print("Nenhuma aba 'Backlog' foi encontrada para consolidar.")