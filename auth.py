import requests
import os
from dotenv import load_dotenv

load_dotenv()  # Carrega variáveis do arquivo .env

# Obter as variáveis de ambiente
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')
RESOURCE = os.getenv('RESOURCE')
SITE_URL=os.getenv('SITE_URL')  

def obter_token_sharepoint():
    
    """Obtém um token de autenticação do SharePoint usando credenciais de cliente.

    Retorna:
        str: O token de acesso do SharePoint ou None se a operação falhar.
    """
    
    url = f"https://accounts.accesscontrol.windows.net/{TENANT_ID}/tokens/OAuth/2"
    
    payload = {
        'grant_type': 'client_credentials',
        'client_id': f"{CLIENT_ID}@{TENANT_ID}",
        'client_secret': CLIENT_SECRET,
        'resource': f"{RESOURCE}@{TENANT_ID}"
    }

    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        return response.json()['access_token']
    except requests.exceptions.RequestException as e:
        print(f"Erro ao obter token: {e}")
        raise Exception("Falha ao obter o token de autenticação.")  # Lança uma exceção

def buscar_listas_sharepoint(token):
    """Busca listas no SharePoint usando o token de autenticação."""
    url = "https://weg365.sharepoint.com/teams/BR-TI-TIN/_api/web/GetFolderByServerRelativeUrl('/teams/BR-TI-TIN/DEV_AlocacaoRecursos/TIN%20-%20Detalhamento%20Atividades')/Files"

    headers = {
        'Content-Type': 'application/json;odata=verbose',
        'Accept': 'application/json;odata=verbose',
        'Authorization': f'Bearer {token}'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida
        return response.json()  # Retorna o JSON da resposta
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")  # Mensagem de erro HTTP
        print(f"Response content: {response.content}")  # Exibe o conteúdo da resposta
        return None
    except requests.exceptions.RequestException as e:
        print(f"Erro ao buscar listas: {e}")
        return None

def buscar_arquivos_pasta(caminho_pasta, token):
    """Busca arquivos em uma pasta específica no SharePoint.

    Args:
        token (str): O token de acesso para autenticação.
        caminho_pasta (str): O caminho da pasta no SharePoint.

    Returns:
        dict: A resposta JSON contendo os arquivos na pasta ou None se a operação falhar.
    """
    url = f"https://weg365.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('{caminho_pasta}')/Files"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()  # Retorna a resposta JSON com os arquivos
    except requests.exceptions.RequestException as e:
        print(f"Erro ao buscar arquivos: {e}")
        return None

def enviar_para_sharepoint(access_token, conteudo_arquivo, nome_destino):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream'  # Use 'application/octet-stream' para enviar arquivos binários
    }

    # Se conteudo_arquivo for uma string, significa que é um caminho de arquivo
    if isinstance(conteudo_arquivo, str):
        with open(conteudo_arquivo, 'rb') as file:
            arquivo_conteudo = file.read()
    else:
        # Assumindo que conteudo_arquivo é um objeto BytesIO
        arquivo_conteudo = conteudo_arquivo.getvalue()

    # Endpoint para upload
    endpoint = f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('/teams/BR-TI-TIN/DEV_AlocacaoRecursos/TIN%20-%20Detalhamento%20Atividades/Consolidado')/Files/add(url='{nome_destino}',overwrite=true)"
    
    response = requests.post(endpoint, headers=headers, data=arquivo_conteudo)

    if response.status_code == 200:
        print("Arquivo enviado com sucesso para o SharePoint.")
    else:
        print("Erro ao enviar o arquivo:", response.json())