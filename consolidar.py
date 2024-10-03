import requests
import os
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv

load_dotenv()

SITE_URL = os.getenv('SITE_URL')

# Função para fazer upload do arquivo para SharePoint
def enviar_para_sharepoint(caminho_arquivo, access_token, nome_destino):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
    }
    
    with open(caminho_arquivo, 'rb') as file:
        arquivo_conteudo = file.read()

    endpoint = f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('/teams/BR-TI-TIN/DEV_AlocacaoRecursos/TIN%20-%20Detalhamento%20Atividades/Consolidado')/Files/add(url='{nome_destino}',overwrite=true)"

    response = requests.post(endpoint, headers=headers, data=arquivo_conteudo)

    if response.status_code == 200:
        print("Arquivo enviado com sucesso para o SharePoint.")
    else:
        print("Erro ao enviar o arquivo:", response.json())

# Função para obter o token do SharePoint
def obter_token_sharepoint():
    url = "https://accounts.accesscontrol.windows.net/886666a6-a8d2-4604-a002-95b622cb7e18/tokens/OAuth/2"
    
    payload = {
        'grant_type': 'client_credentials',
        'client_id': 'YOUR_CLIENT_ID',
        'client_secret': 'YOUR_CLIENT_SECRET',
        'resource': '00000003-0000-0ff1-ce00-000000000000/weg365.sharepoint.com@886666a6-a8d2-4604-a002-95b622cb7e18'
    }

    response = requests.post(url, data=payload)

    if response.status_code == 200:
        return response.json()['access_token']
    else:
        print("Erro ao obter token:", response.text)
        return None

# Função para baixar arquivos do SharePoint
def baixar_arquivo_sharepoint(url_arquivo, token):
    headers = {
        'Authorization': f'Bearer {token}'
    }
    
    response = requests.get(url_arquivo, headers=headers)
    
    if response.status_code == 200:
        return BytesIO(response.content)
    else:
        print(f"Erro ao baixar arquivo {url_arquivo}: {response.status_code}")
        return None

# Função principal para consolidar planilhas do SharePoint
def consolidar_planilhas_sharepoint(lista_arquivos_sharepoint):
    """Consolida todas as planilhas dos arquivos Excel do SharePoint, ignorando a aba 'Backlog'."""
    token = obter_token_sharepoint()
    
    if not token:
        print("Erro ao obter token. Não foi possível continuar.")
        return
    
    lista_dfs = []
    
    # Loop para percorrer todos os arquivos no SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")
        
        # Usa requests para acessar diretamente o arquivo no SharePoint
        response = requests.get(url_arquivo, headers={'Authorization': f'Bearer {token}'})
        
        if response.status_code != 200:
            print(f"Erro ao acessar arquivo: {url_arquivo} - {response.status_code}")
            continue
        
        # Usa o conteúdo diretamente, sem baixar
        arquivo_excel = BytesIO(response.content)  # Lê o conteúdo da resposta no formato Excel
        
        xls = pd.ExcelFile(arquivo_excel)
        
        for nome_aba in xls.sheet_names:
            # Verifica se a aba é "Backlog" e ignora
            if 'Backlog' in nome_aba:
                print(f"Ignorando aba: '{nome_aba}'")
                continue

            df = pd.read_excel(xls, sheet_name=nome_aba)

            # Verifica se a aba tem as colunas mínimas necessárias
            if 'Planned effort' in df.columns and df.shape[1] > 5:
                valor_secao = df.iloc[3, 0]  # Captura a célula A4
                valor_equipe = df.iloc[4, 0]  # Captura a célula A5

                for index, row in df.iterrows():
                    if index < 3:  # Ignora as primeiras linhas
                        continue

                    epic = row['Epic'] if 'Epic' in df.columns else ''
                    status = row['Status'] if 'Status' in df.columns else ''
                    due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                    planned_effort = row['Planned effort']
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
                                'Due Date': due_date,
                                'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                                'Planned Effort': planned_effort,
                                'Estimate Effort': estimate_effort,
                                'MÊS': mes_abreviado,
                                'ANO': ano,
                                'Horas mês': valor_hora_mes,
                                'Seção': valor_secao,
                                'Equipe': valor_equipe
                            }
                            lista_dfs.append(nova_linha)

                        except Exception as e:
                            print(f"Erro no arquivo '{url_arquivo}', aba '{nome_aba}', linha {index + 1}, coluna '{mes}': {str(e)}")

    # Consolida os dados em um DataFrame
    if lista_dfs:
        # Consolida os dados em um DataFrame
        dataframe_consolidado = pd.DataFrame(lista_dfs)
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_excel(arquivo_memoria, index=False, engine='openpyxl')
        arquivo_memoria.seek(0)  # Move o ponteiro de volta ao início do arquivo

        # URL do diretório no SharePoint onde você deseja salvar o arquivo
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        
        # Nome do arquivo a ser salvo no SharePoint
        nome_arquivo_sharepoint = "planilha_consolidada.xlsx"

        try:
            # Faz o upload do arquivo diretamente da memória para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint, url_diretorio_sharepoint)
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")


# Função principal para consolidar a aba "Backlog" e enviar para o SharePoint
def consolidar_aba_backlog_sharepoint(lista_arquivos_sharepoint):
    """Consolida a aba 'Backlog' de todos os arquivos Excel no SharePoint."""
    token = obter_token_sharepoint()
    
    if not token:
        print("Erro ao obter token. Não foi possível continuar.")
        return
    
    lista_dfs = []
    
    # Itera sobre todos os arquivos no SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")
        arquivo_excel = baixar_arquivo_sharepoint(url_arquivo, token)
        
        if not arquivo_excel:
            continue
        
        xls = pd.ExcelFile(arquivo_excel)

        # Processa a aba "Backlog"
        for nome_aba in xls.sheet_names:
            if 'Backlog' in nome_aba:
                print(f"Processando aba: '{nome_aba}'")

                # Lê a aba "Backlog" como DataFrame
                df = pd.read_excel(xls, sheet_name=nome_aba)

                if 'Estimated effort' in df.columns and df.shape[1] > 5:
                    valor_secao = df.iloc[3, 0]
                    valor_equipe = df.iloc[4, 0]

                    for index, row in df.iterrows():
                        if index < 5:
                            continue

                        epic = row['Epic'] if 'Epic' in df.columns else ''
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        estimated_effort = row['Estimated effort']

                        nova_linha = {
                            'Epic': epic,
                            'Status': status,
                            'Due Date': due_date,
                            'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                            'Estimated Effort': estimated_effort,
                            'Seção': valor_secao,
                            'Equipe': valor_equipe
                        }
                        lista_dfs.append(nova_linha)

    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')
        dataframe_consolidado = dataframe_consolidado[dataframe_consolidado.iloc[:, 0].notna()]

        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_excel(arquivo_memoria, index=False, engine='openpyxl')
        arquivo_memoria.seek(0)  # Move o ponteiro de volta ao início do arquivo

        # URL do diretório no SharePoint onde você deseja salvar o arquivo
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        
        # Nome do arquivo a ser salvo no SharePoint
        nome_arquivo_sharepoint = "backlog_consolidado.xlsx"

        try:
            # Faz o upload do arquivo diretamente da memória para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint, url_diretorio_sharepoint)
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")
    else:
        print("Nenhuma aba 'Backlog' foi consolidada. Verifique os arquivos de entrada.")

def consolidar_horas_backlog_sharepoint(lista_arquivos_sharepoint):
    """Consolida os dados das abas 'Backlog' de todos os arquivos Excel no SharePoint."""
    token = obter_token_sharepoint()
    
    if not token:
        print("Erro ao obter token. Não foi possível continuar.")
        return
    
    lista_dfs = []

    # Itera sobre todos os arquivos no SharePoint
    for url_arquivo in lista_arquivos_sharepoint:
        print(f"Processando arquivo: {url_arquivo}")
        arquivo_excel = baixar_arquivo_sharepoint(url_arquivo, token)
        
        if not arquivo_excel:
            continue
        
        xls = pd.ExcelFile(arquivo_excel)

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
                    horas_mes = row[7:19]  # Colunas H até S

                    # Verifica se as células contêm dados válidos
                    if not horas_mes.isnull().all():
                        # Cria uma linha para cada mês, com o respectivo valor de Hora/mês
                        for i, hora in enumerate(horas_mes):
                            nova_linha = {
                                'Epic': epic,
                                'Planned effort': planned_effort,
                                'Hora/mês': hora,
                                'Mês': f'Mês {i + 1}'
                            }
                            lista_dfs.append(nova_linha)

    # Consolida os dados em um DataFrame final
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Salva o DataFrame em um objeto de memória (BytesIO), sem salvar localmente
        arquivo_memoria = BytesIO()
        dataframe_consolidado.to_excel(arquivo_memoria, index=False, engine='openpyxl')
        arquivo_memoria.seek(0)  # Move o ponteiro de volta ao início do arquivo

        # URL do diretório no SharePoint onde você deseja salvar o arquivo
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        
        # Nome do arquivo a ser salvo no SharePoint
        nome_arquivo_sharepoint = "planilha_consolidada_backlog_horas.xlsx"

        try:
            # Faz o upload do DataFrame consolidado diretamente para o SharePoint
            enviar_para_sharepoint(token, arquivo_memoria, nome_arquivo_sharepoint, url_diretorio_sharepoint)
            print(f"Relatório consolidado enviado para o SharePoint com sucesso.")
        except Exception as e:
            print(f"Erro ao enviar o arquivo '{nome_arquivo_sharepoint}' para o SharePoint: {str(e)}")
    else:
        print("Nenhuma aba 'Backlog' foi encontrada para consolidar.")

