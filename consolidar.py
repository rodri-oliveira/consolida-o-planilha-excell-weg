import requests
import os
import pandas as pd
from io import BytesIO

# Função para fazer upload do arquivo para SharePoint
def upload_arquivo_sharepoint(token, caminho_arquivo, nome_arquivo, url_diretorio):
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream'
    }

    with open(caminho_arquivo, 'rb') as arquivo:
        response = requests.put(url_diretorio + f"/{nome_arquivo}:/content", headers=headers, data=arquivo)

    if response.status_code == 201:
        print(f"Arquivo '{nome_arquivo}' enviado com sucesso para o SharePoint.")
    else:
        print(f"Erro ao enviar arquivo '{nome_arquivo}': {response.text}")

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
        arquivo_excel = baixar_arquivo_sharepoint(url_arquivo, token)
        
        if not arquivo_excel:
            continue
        
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
        dataframe_consolidado = pd.DataFrame(lista_dfs)
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        caminho_para_salvar_arquivo = 'C:/consolida-o-planilha-excell-weg/planilhas-consolidada/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        # URL do diretório no SharePoint onde você deseja salvar o arquivo
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        
        # Nome do arquivo a ser salvo no SharePoint
        nome_arquivo_sharepoint = "planilha_consolidada.xlsx"

        # Faz o upload do arquivo para o SharePoint
        upload_arquivo_sharepoint(token, caminho_para_salvar_arquivo, nome_arquivo_sharepoint, url_diretorio_sharepoint)
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
        
        for nome_aba in xls.sheet_names:
            if 'Backlog' not in nome_aba:
                continue
            
            df = pd.read_excel(xls, sheet_name=nome_aba)
            
            # Verifica se a aba tem as colunas mínimas necessárias
            if 'Backlog' in df.columns and df.shape[1] > 5:
                for index, row in df.iterrows():
                    if index < 3:  # Ignora as primeiras linhas
                        continue
                    
                    # Continue a mesma lógica como nas outras funções...

# Chame a função principal conforme necessário
# lista_arquivos = ['url_do_arquivo_1', 'url_do_arquivo_2', ...]
# consolidar_planilhas_sharepoint(lista_arquivos)

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

        # Faz o upload do DataFrame consolidado diretamente para o SharePoint
        caminho_para_salvar_arquivo = 'planilha_consolidada_backlog_horas.xlsx'
        url_diretorio_sharepoint = "https://weg365.sharepoint.com/sites/BR-TI-TIN/DEV_AlocacaoRecursos/Consolidado"
        
        # Faz o upload da planilha consolidada para o SharePoint
        upload_arquivo_sharepoint(token, caminho_para_salvar_arquivo, caminho_para_salvar_arquivo, url_diretorio_sharepoint)
        
        print(f"Relatório consolidado enviado para o SharePoint com sucesso.")
    else:
        print("Nenhuma aba 'Backlog' foi encontrada para consolidar.")

