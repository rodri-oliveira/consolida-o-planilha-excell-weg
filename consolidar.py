import pandas as pd
import os

def ajustar_ano(meses):
    anos = []
    for mes in meses:
        mes_valor, ano_valor = mes.split('/')
        
        # Mapeia os meses em português para seus respectivos números
        meses_portugues = {
            'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12,
            'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
            'jul': 7, 'ago': 8
        }
        
        mes_num = meses_portugues[mes_valor]  # Converte o mês para número
        
        # Lógica para determinar o ano
        if mes_num <= 5:  # De ago (8) até dez (12) de 2024
            ano = 2024
        elif mes_num <= 17:  # De jan (1) até dez (12) de 2025
            ano = 2025
        elif mes_num <= 29:  # De jan (1) até dez (12) de 2026
            ano = 2026
        else:  # Para meses além de jan/26
            ano = 2024 + (mes_num - 1) // 12
        
        anos.append(ano)
    return anos

def consolidar_planilhas(caminho_das_planilhas):
    lista_dfs = []
    
    # Lista de meses
    meses = [
        'ago/24', 'set/24', 'out/24', 'nov/24', 'dez/24',  # 2024
        'jan/25', 'fev/25', 'mar/25', 'abr/25', 'mai/25', 'jun/25', 
        'jul/25', 'ago/25', 'set/25', 'out/25', 'nov/25', 'dez/25',  # 2025
        'jan/26', 'fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
        'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26'  # 2026
        # Adicione mais meses conforme necessário
    ]

    anos = ajustar_ano(meses)  # Chama a função para obter os anos

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            for nome_aba in xls.sheet_names:
                if nome_aba == "Backlog":
                    continue

                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verificando se as colunas necessárias existem e se a planilha tem mais que 5 colunas
                if 'Planned effort' in df.columns and df.shape[1] > 5:
                    for index, row in df.iterrows():
                        if index < 4:
                            continue

                        epic = df.iloc[1, 0]  # Valor de Epic na linha 2
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        planned_effort = row['Planned effort']
                        estimate_effort = row.iloc[5] if len(row) > 5 else None

                        # Identificando as colunas de meses dinamicamente
                        colunas_meses = [col for col in df.columns if any(mes in col for mes in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])]

                        for idx, mes in enumerate(colunas_meses):
                            valor_hora_mes = row[mes]
                            
                            # Usa o ano correspondente da lista obtida
                            ano = anos[idx] if idx < len(anos) else ''
                            
                            nova_linha = {
                                'Epic': epic,
                                'Status': status,
                                'Due Date': due_date,
                                'Assignee': row['Assignee'] if 'Assignee' in df.columns else '',
                                'Planned Effort': planned_effort,
                                'Estimate Effort': estimate_effort,
                                'MÊS': meses[idx] if idx < len(meses) else '',  # Preenche com o mês correspondente
                                'ANO': ano,
                                'Horas mês': valor_hora_mes
                            }
                            lista_dfs.append(nova_linha)

    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Exclui colunas indesejadas
        dataframe_consolidado.drop(columns=['Horas disponíveis', 'Total de esforço'], inplace=True, errors='ignore')

        caminho_para_salvar_arquivo = 'C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")

# Chamada da função com o caminho das planilhas
caminho_das_planilhas = 'C:/consolidar-planilha-weg/planilhas-base/'
consolidar_planilhas(caminho_das_planilhas)
