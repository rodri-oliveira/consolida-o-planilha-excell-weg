import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# Variável global para armazenar o DataFrame consolidado
dataframe = None

# Função responsável por consolidar as planilhas
def consolidar_planilhas(caminho_pasta, caminho_saida):
    global dataframe  # Define o DataFrame como global para ser usado em outras funções
    lista_dataframes = []  # Lista que armazenará os DataFrames de cada planilha

    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(caminho_pasta):
        # Verifica se o arquivo é um Excel
        if nome_arquivo.endswith('.xlsx'):
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)  # Caminho completo do arquivo
            xls = pd.ExcelFile(caminho_arquivo_excel)  # Carrega o arquivo Excel

            # Itera sobre todas as abas do arquivo Excel
            for sheet_name in xls.sheet_names:
                # Lê a aba específica em um DataFrame
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)
                lista_dataframes.append(dataframe_planilha)  # Adiciona o DataFrame à lista

    # Se houver planilhas carregadas
    if lista_dataframes:
        # Concatena todas as planilhas em um único DataFrame
        dataframe = pd.concat(lista_dataframes, ignore_index=True)
        # Salva o DataFrame consolidado em um arquivo Excel
        dataframe.to_excel(caminho_saida, index=False)
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
    else:
        # Caso nenhuma planilha tenha sido encontrada
        print("Nenhuma planilha foi encontrada ou consolidada.")

# Função para gerar um relatório filtrado com base nos critérios
def gerar_relatorio(dataframe, coluna, operador, valor):
    # Aplica o filtro com base no operador escolhido
    if operador == "maior que":
        return dataframe[dataframe[coluna] > valor]
    elif operador == "menor que":
        return dataframe[dataframe[coluna] < valor]
    elif operador == "igual a":
        return dataframe[dataframe[coluna] == valor]
    return pd.DataFrame()  # Retorna um DataFrame vazio caso não haja correspondência

# Função para configurar a interface gráfica do usuário
def consolidar_planilhas_interface():
    # Função para consolidar as planilhas
    def consolidar():
        caminho_pasta = entrada_pasta.get()  # Obtém o caminho da pasta digitado pelo usuário
        caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx"  # Define o caminho de saída
        consolidar_planilhas(caminho_pasta, caminho_saida)  # Chama a função de consolidação

        # Atualiza o combobox com os nomes das colunas
        if dataframe is not None:
            coluna_combobox['values'] = dataframe.columns.tolist()  # Preenche o combobox com os nomes das colunas

        messagebox.showinfo("Sucesso", "Planilhas consolidadas com sucesso!")  # Mostra mensagem de sucesso

    # Função para gerar o relatório baseado nos filtros
    def gerar_relatorio_callback():
        coluna = coluna_combobox.get()  # Obtém a coluna selecionada
        operador = operador_combobox.get()  # Obtém o operador selecionado
        valor = valor_entry.get()  # Obtém o valor inserido

        # Verifica se todos os campos foram preenchidos
        if coluna and operador and valor:
            try:
                valor = float(valor)  # Converte o valor inserido para float
                relatorio = gerar_relatorio(dataframe, coluna, operador, valor)  # Gera o relatório

                # Verifica se o relatório está vazio
                if relatorio.empty:
                    messagebox.showinfo("Resultado", "Nenhum dado encontrado para os critérios selecionados.")
                else:
                    caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/relatorio.xlsx"  # Caminho do relatório gerado
                    relatorio.to_excel(caminho_saida, index=False)  # Salva o relatório
                    messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em {caminho_saida}!")
            except ValueError:
                messagebox.showwarning("Erro", "Insira um valor numérico válido.")  # Mensagem de erro para valor inválido
        else:
            messagebox.showwarning("Atenção", "Preencha todos os campos antes de gerar o relatório.")  # Aviso para campos vazios

    # Função para buscar a pasta de planilhas
    def buscar_pasta():
        caminho = filedialog.askdirectory()  # Abre o explorador de arquivos para selecionar a pasta
        if caminho:
            entrada_pasta.delete(0, tk.END)  # Limpa o campo de entrada
            entrada_pasta.insert(0, caminho)  # Insere o caminho selecionado

    # Função para limpar os campos da interface (Nova Busca)
    def nova_busca():
        entrada_pasta.delete(0, tk.END)  # Limpa o campo de caminho da pasta
        coluna_combobox.set('')  # Limpa o combobox de colunas
        operador_combobox.set('')  # Limpa o combobox de operadores
        valor_entry.delete(0, tk.END)  # Limpa o campo de valor

    # Criação da janela principal
    janela = tk.Tk()
    janela.title("Consolidar Planilhas")  # Define o título da janela

    # Campo para inserir o caminho da pasta
    tk.Label(janela, text="Caminho da pasta:").pack()  # Label para o caminho da pasta
    entrada_pasta = tk.Entry(janela, width=50)  # Campo de texto para inserir o caminho
    entrada_pasta.pack()

    # Botão para buscar a pasta
    botao_buscar = tk.Button(janela, text="Buscar Pasta", command=buscar_pasta)
    botao_buscar.pack()

    # Botão para consolidar as planilhas
    botao_consolidar = tk.Button(janela, text="Consolidar", command=consolidar)
    botao_consolidar.pack()

    # Seção de filtros para gerar o relatório
    tk.Label(janela, text="Coluna para o relatório:").pack()  # Label para seleção de coluna
    coluna_combobox = ttk.Combobox(janela, values=[])  # Combobox para escolher a coluna
    coluna_combobox.pack()

    tk.Label(janela, text="Operador:").pack()  # Label para o operador
    operador_combobox = ttk.Combobox(janela, values=["maior que", "menor que", "igual a"])  # Combobox para escolher o operador
    operador_combobox.pack()

    tk.Label(janela, text="Valor:").pack()  # Label para o valor
    valor_entry = tk.Entry(janela)  # Campo de texto para inserir o valor
    valor_entry.pack()

    # Botão para gerar o relatório
    botao_gerar_relatorio = tk.Button(janela, text="Gerar Relatório", command=gerar_relatorio_callback)
    botao_gerar_relatorio.pack()

    # Botão "Nova Busca" para limpar os campos
    botao_nova_busca = tk.Button(janela, text="Nova Busca", command=nova_busca)
    botao_nova_busca.pack()

    # Inicia o loop da interface gráfica
    janela.mainloop()
