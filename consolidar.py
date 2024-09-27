import pandas as pd
import os

def consolidar_planilhas(caminho_das_planilhas):
    """Consolida todas as planilhas de todos os arquivos Excel no diretório especificado."""
    global dataframe_consolidado
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            # Processa todas as abas dinamicamente, ignorando a aba "Backlog"
            for nome_aba in xls.sheet_names:
                if nome_aba == "Backlog":  # Ignora a aba "Backlog"
                    continue

                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verifica se a aba tem as colunas mínimas necessárias para a consolidação
                if 'Planned effort' in df.columns and df.shape[1] > 5:
                    for index, row in df.iterrows():
                        if index < 4:
                            continue  # Ignora as primeiras linhas (ajuste conforme necessário)

                        # Obtém os valores essenciais (Epic, Status, Due Date, etc.)
                        epic = df.iloc[1, 0]  # Captura o valor de 'Epic' na linha 2
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        planned_effort = row['Planned effort']
                        estimate_effort = row.iloc[5] if len(row) > 5 else None

                        # Identifica as colunas de meses dinamicamente com base em padrões de nomes de meses
                        colunas_meses = [col for col in df.columns if any(mes in col for mes in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])]

                        # Definição dinâmica de ano inicial e cálculo de ano conforme a progressão dos meses
                        ano_base = 2024
                        meses_por_ano = 12

                        # Processa dinamicamente as colunas de meses
                        for idx, mes in enumerate(colunas_meses):
                            valor_hora_mes = row[mes]
                            ano = ano_base + (idx // meses_por_ano)  # Calcula o ano automaticamente baseado no número de meses

                            # Verifica se o valor da célula é válido e se é numérico
                            if pd.isnull(valor_hora_mes) or not isinstance(valor_hora_mes, (int, float)):
                                continue

                            # Cria uma nova linha no DataFrame consolidado
                            nova_linha = {
                                'Epic': epic,
                                'Status': status,
                                'Due Date': due_date,
                                'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                                'Planned Effort': planned_effort,
                                'Estimate Effort': estimate_effort,
                                'MÊS': mes,
                                'ANO': ano,
                                'Horas mês': valor_hora_mes
                            }

                            lista_dfs.append(nova_linha)

    # Consolida todos os dados das abas em um único DataFrame
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Remove colunas indesejadas
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        # Define o caminho para salvar a planilha consolidada
        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")
