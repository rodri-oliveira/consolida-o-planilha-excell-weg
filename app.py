import os  # Módulo para interagir com o sistema de arquivos
import pandas as pd  # Biblioteca para manipulação de dados em DataFrames
import tkinter as tk  # Biblioteca para criar interfaces gráficas
from tkinter import messagebox, ttk, filedialog  # Importando componentes do tkinter

# Variável global para armazenar o DataFrame consolidado
dataframe = None

def consolidar_planilhas(caminho_pasta, caminho_saida):
    global dataframe  # Usando a variável global dataframe
    lista_dataframes = []  # Lista para armazenar os DataFrames de cada planilha

    # Itera sobre os arquivos na pasta especificada
    for nome_arquivo in os.listdir(caminho_pasta):
        if nome_arquivo.endswith('.xlsx'):  # Verifica se é um arquivo Excel
            caminho_arquivo_excel = os.path.join(caminho_pasta, nome_arquivo)  # Caminho completo do arquivo
            xls = pd.ExcelFile(caminho_arquivo_excel)  # Carrega o arquivo Excel

            # Itera sobre as abas do arquivo
            for sheet_name in xls.sheet_names:
                dataframe_planilha = pd.read_excel(caminho_arquivo_excel, sheet_name=sheet_name)  # Lê a aba
                lista_dataframes.append(dataframe_planilha)  # Adiciona o DataFrame à lista

    # Se houver DataFrames na lista, combina-os
    if lista_dataframes:
        dataframe = pd.concat(lista_dataframes, ignore_index=True)  # Combina todos os DataFrames
        dataframe.to_excel(caminho_saida, index=False)  # Salva o DataFrame consolidado em um novo arquivo
        print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
    else:
        print("Nenhuma planilha foi encontrada ou consolidada.")  # Mensagem de erro se não houver planilhas

def gerar_relatorio(dataframe, coluna, operador, valor):
    # Filtra o DataFrame com base no operador escolhido
    if operador == "maior que":
        return dataframe[dataframe[coluna] > valor]  # Filtra valores maiores
    elif operador == "menor que":
        return dataframe[dataframe[coluna] < valor]  # Filtra valores menores
    elif operador == "igual a":
        return dataframe[dataframe[coluna] == valor]  # Filtra valores iguais
    return pd.DataFrame()  # Retorna um DataFrame vazio se nada for encontrado

def consolidar_planilhas_interface():
    # Função para a interface gráfica
    def consolidar():
        caminho_pasta = entrada_pasta.get()  # Obtém o caminho da pasta a partir da entrada
        caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx"
        consolidar_planilhas(caminho_pasta, caminho_saida)  # Chama a função de consolidação

        # Atualiza o combobox com os nomes das colunas
        if dataframe is not None:
            coluna_combobox['values'] = dataframe.columns.tolist()  # Adiciona as colunas ao combobox

        messagebox.showinfo("Sucesso", "Planilhas consolidadas com sucesso!")  # Mensagem de sucesso

    def gerar_relatorio_callback():
        coluna = coluna_combobox.get()  # Obtém a coluna selecionada
        operador = operador_combobox.get()  # Obtém o operador selecionado
        valor = valor_entry.get()  # Obtém o valor da entrada
        
        if coluna and operador and valor:  # Verifica se todos os campos foram preenchidos
            try:
                valor = float(valor)  # Converte o valor para float
                relatorio = gerar_relatorio(dataframe, coluna, operador, valor)  # Gera o relatório

                if relatorio.empty:  # Verifica se o relatório está vazio
                    messagebox.showinfo("Resultado", "Nenhum dado encontrado para os critérios selecionados.")
                else:
                    caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/relatorio.xlsx"
                    relatorio.to_excel(caminho_saida, index=False)  # Salva o relatório em um novo arquivo
                    messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em {caminho_saida}!")  # Mensagem de sucesso
            except ValueError:
                messagebox.showwarning("Erro", "Insira um valor numérico válido.")  # Mensagem de erro se a conversão falhar
        else:
            messagebox.showwarning("Atenção", "Preencha todos os campos antes de gerar o relatório.")  # Aviso se campos não foram preenchidos

    def buscar_pasta():
        caminho = filedialog.askdirectory()  # Abre o diálogo para selecionar uma pasta
        if caminho:
            entrada_pasta.delete(0, tk.END)  # Limpa o campo de entrada
            entrada_pasta.insert(0, caminho)  # Insere o caminho selecionado

    # Criação da janela principal
    janela = tk.Tk()
    janela.title("Consolidar Planilhas")  # Título da janela

    tk.Label(janela, text="Caminho da pasta:").pack()  # Label para o caminho da pasta
    entrada_pasta = tk.Entry(janela, width=50)  # Campo de entrada para o caminho
    entrada_pasta.pack()

    botao_buscar = tk.Button(janela, text="Buscar Pasta", command=buscar_pasta)  # Botão para buscar a pasta
    botao_buscar.pack()

    botao_consolidar = tk.Button(janela, text="Consolidar", command=consolidar)  # Botão para consolidar planilhas
    botao_consolidar.pack()

    tk.Label(janela, text="Coluna para o relatório:").pack()  # Label para a coluna do relatório
    coluna_combobox = ttk.Combobox(janela, values=[])  # Combobox para selecionar a coluna
    coluna_combobox.pack()

    tk.Label(janela, text="Operador:").pack()  # Label para o operador
    operador_combobox = ttk.Combobox(janela, values=["maior que", "menor que", "igual a"])  # Combobox para operadores
    operador_combobox.pack()

    tk.Label(janela, text="Valor:").pack()  # Label para o valor
    valor_entry = tk.Entry(janela)  # Campo de entrada para o valor
    valor_entry.pack()

    botao_gerar_relatorio = tk.Button(janela, text="Gerar Relatório", command=gerar_relatorio_callback)  # Botão para gerar relatório
    botao_gerar_relatorio.pack()

    janela.mainloop()  # Inicia o loop da interface gráfica
