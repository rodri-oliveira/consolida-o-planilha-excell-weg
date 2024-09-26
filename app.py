# app.py
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from consolidar import consolidar_planilhas  # Importa a função consolidar_planilhas do arquivo consolidar.py

# Variável global para armazenar o DataFrame consolidado
dataframe_consolidado = None

def gerar_relatorio(dataframe, nome_coluna, operador, valor):
    """Filtra o DataFrame baseado nos critérios fornecidos."""
    if operador == "maior que":
        return dataframe[dataframe[nome_coluna] > valor]
    elif operador == "menor que":
        return dataframe[dataframe[nome_coluna] < valor]
    elif operador == "igual a":
        return dataframe[dataframe[nome_coluna] == valor]
    else:
        raise ValueError("Operador desconhecido.")

def consolidar_planilhas_interface():
    global dataframe_consolidado

    def consolidar():
        caminho_pasta_origem = entrada_pasta.get()
        consolidar_planilhas(caminho_pasta_origem)

        if dataframe_consolidado is not None:
            coluna_combobox['values'] = dataframe_consolidado.columns.tolist()

        messagebox.showinfo("Sucesso", "Planilhas consolidadas com sucesso!")

    def gerar_relatorio_callback():
        nome_coluna = coluna_combobox.get()
        operador = operador_combobox.get()
        valor = valor_entry.get()

        if nome_coluna and operador and valor:
            try:
                valor = float(valor)
                relatorio_filtrado = gerar_relatorio(dataframe_consolidado, nome_coluna, operador, valor)

                if relatorio_filtrado.empty:
                    messagebox.showinfo("Resultado", "Nenhum dado encontrado para os critérios selecionados.")
                else:
                    caminho_saida_relatorio = "C:/consolidar-planilha-weg/relatorios/relatorio.xlsx"
                    os.makedirs(os.path.dirname(caminho_saida_relatorio), exist_ok=True)

                    relatorio_filtrado.to_excel(caminho_saida_relatorio, index=False)
                    messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em {caminho_saida_relatorio}!")
            except ValueError:
                messagebox.showwarning("Erro", "Insira um valor numérico válido.")
        else:
            messagebox.showwarning("Atenção", "Preencha todos os campos antes de gerar o relatório.")

    def buscar_pasta():
        caminho_selecionado = filedialog.askdirectory()
        if caminho_selecionado:
            entrada_pasta.delete(0, tk.END)
            entrada_pasta.insert(0, caminho_selecionado)

    def nova_busca():
        entrada_pasta.delete(0, tk.END)
        coluna_combobox.set('')
        operador_combobox.set('')
        valor_entry.delete(0, tk.END)

    # Criação da janela principal
    janela_principal = tk.Tk()
    janela_principal.title("Consolidar Planilhas")

    # Interface de seleção de pasta e comandos
    tk.Label(janela_principal, text="Caminho da pasta:").pack()
    entrada_pasta = tk.Entry(janela_principal, width=50)
    entrada_pasta.pack()

    botao_buscar_pasta = tk.Button(janela_principal, text="Buscar Pasta", command=buscar_pasta)
    botao_buscar_pasta.pack()

    botao_consolidar = tk.Button(janela_principal, text="Consolidar", command=consolidar)
    botao_consolidar.pack()

    # Seletor de colunas, operadores e valor para filtrar o relatório
    tk.Label(janela_principal, text="Selecione uma coluna:").pack()
    coluna_combobox = ttk.Combobox(janela_principal)
    coluna_combobox.pack()

    tk.Label(janela_principal, text="Selecione um operador:").pack()
    operador_combobox = ttk.Combobox(janela_principal, values=["maior que", "menor que", "igual a"])
    operador_combobox.pack()

    tk.Label(janela_principal, text="Valor:").pack()
    valor_entry = tk.Entry(janela_principal)
    valor_entry.pack()

    # Botões para gerar relatório e limpar a busca
    botao_gerar_relatorio = tk.Button(janela_principal, text="Gerar Relatório", command=gerar_relatorio_callback)
    botao_gerar_relatorio.pack()

    botao_nova_busca = tk.Button(janela_principal, text="Nova Busca", command=nova_busca)
    botao_nova_busca.pack()

    # Inicia a interface
    janela_principal.mainloop()

# Função para consolidar as planilhas
def consolidar_planilhas(caminho_das_planilhas):
    """Consolida todas as planilhas de todos os arquivos Excel no diretório especificado."""
    global dataframe_consolidado
    lista_dfs = []

    # Loop para percorrer todos os arquivos no diretório de planilhas
    for arquivo in os.listdir(caminho_das_planilhas):
        if arquivo.endswith('.xlsx'):
            caminho_completo = os.path.join(caminho_das_planilhas, arquivo)
            xls = pd.ExcelFile(caminho_completo)

            # Itera sobre todas as abas do arquivo Excel
            for nome_aba in xls.sheet_names:
                if nome_aba == "Backlog":
                    print(f"Aba '{nome_aba}' do arquivo {arquivo} foi ignorada.")
                    continue

                df = pd.read_excel(xls, sheet_name=nome_aba)

                # Verifica se o DataFrame contém as colunas necessárias
                if 'Planned effort' in df.columns and df.shape[1] > 5:
                    for index, row in df.iterrows():
                        if index < 0:  # Ignora as primeiras 4 linhas
                            continue

                        # Captura os dados relevantes
                        epic = row['Epic'] if 'Epic' in df.columns else ''
                        status = row['Status'] if 'Status' in df.columns else ''
                        due_date = row['Due Date'] if 'Due Date' in df.columns else ''
                        planned_effort = row['Planned effort']
                        estimate_effort = row.iloc[5] if len(row) > 5 else None

                        # Verifique se as colunas de I a Y estão presentes no DataFrame
                        colunas_meses = df.columns[8:25]

                        for idx, mes in enumerate(colunas_meses):
                            valor_hora_mes = row[mes]
                            ano = 2024 if idx < 5 else 2025

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

    # Cria um DataFrame a partir da lista de dicionários
    if lista_dfs:
        dataframe_consolidado = pd.DataFrame(lista_dfs)

        # Exclui colunas indesejadas, mas mantém as linhas de "Atividades DTI_Suporte"
        dataframe_consolidado.drop(columns=['GAP', 'Horas disponíveis', 'Total de esforço (hrs)'], inplace=True, errors='ignore')

        # Salvando o DataFrame consolidado no caminho especificado
        caminho_para_salvar_arquivo = 'C:/consolida-o-planilha-excell-weg/planilhas-consolidadas/planilha_consolidada.xlsx'
        os.makedirs(os.path.dirname(caminho_para_salvar_arquivo), exist_ok=True)
        dataframe_consolidado.to_excel(caminho_para_salvar_arquivo, index=False)

        print(f"Relatório consolidado salvo em: {caminho_para_salvar_arquivo}")
    else:
        print("Nenhuma planilha foi consolidada. Verifique os arquivos de entrada.")
