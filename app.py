# Importa as bibliotecas necessárias
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# Variável global para armazenar o DataFrame consolidado
dataframe_consolidado = None

def consolidar_planilhas(caminho_pasta_origem):
    global dataframe_consolidado
    lista_dataframes = []

    try:
        # Itera sobre todos os arquivos na pasta de origem
        for nome_arquivo in os.listdir(caminho_pasta_origem):
            if nome_arquivo.endswith('.xlsx'):
                caminho_arquivo_excel = os.path.join(caminho_pasta_origem, nome_arquivo)
                xls = pd.ExcelFile(caminho_arquivo_excel)

                # Itera sobre todas as abas de cada arquivo Excel
                for nome_aba in xls.sheet_names:
                    dataframe_aba = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba)
                    lista_dataframes.append(dataframe_aba)

        if lista_dataframes:
            dataframe_consolidado = pd.concat(lista_dataframes, ignore_index=True)
            caminho_saida = "C:/consolidar-planilha-weg/planilhas-consolidadas/planilha_consolidada.xlsx"
            os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
            dataframe_consolidado.to_excel(caminho_saida, index=False)
            print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
        else:
            print("Nenhuma planilha foi encontrada ou consolidada.")
    
    except FileNotFoundError:
        messagebox.showwarning("Erro", f"O caminho '{caminho_pasta_origem}' não foi encontrado.")
    except PermissionError:
        messagebox.showwarning("Erro", f"Você não tem permissão para acessar o arquivo: '{caminho_arquivo_excel}'.")
    except pd.errors.EmptyDataError:
        messagebox.showwarning("Erro", "Um dos arquivos Excel está vazio.")
    except Exception as e:
        messagebox.showwarning("Erro", f"Ocorreu um erro: {str(e)}")

def gerar_relatorio(dataframe, nome_coluna, operador, valor):
    try:
        if nome_coluna not in dataframe.columns:
            messagebox.showwarning("Erro", f"A coluna '{nome_coluna}' não existe no DataFrame.")
            return pd.DataFrame()

        dataframe[nome_coluna] = pd.to_numeric(dataframe[nome_coluna], errors='coerce')

        if operador == "maior que":
            return dataframe[dataframe[nome_coluna] > valor]
        elif operador == "menor que":
            return dataframe[dataframe[nome_coluna] < valor]
        elif operador == "igual a":
            return dataframe[dataframe[nome_coluna] == valor]
        return pd.DataFrame()
    except Exception as e:
        messagebox.showwarning("Erro", f"Ocorreu um erro ao aplicar o filtro: {str(e)}")
        return pd.DataFrame()

def consolidar_planilhas_interface():
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
                    caminho_saida_relatorio = "C:/consolida-o-planilha-excell-weg/relatórios/relatorio.xlsx"
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

    janela_principal = tk.Tk()
    janela_principal.title("Consolidar Planilhas")

    tk.Label(janela_principal, text="Caminho da pasta:").pack()
    entrada_pasta = tk.Entry(janela_principal, width=50)
    entrada_pasta.pack()

    botao_buscar_pasta = tk.Button(janela_principal, text="Buscar Pasta", command=buscar_pasta)
    botao_buscar_pasta.pack()

    botao_consolidar = tk.Button(janela_principal, text="Consolidar", command=consolidar)
    botao_consolidar.pack()

    tk.Label(janela_principal, text="Selecione uma coluna:").pack()
    coluna_combobox = ttk.Combobox(janela_principal)
    coluna_combobox.pack()

    tk.Label(janela_principal, text="Selecione um operador:").pack()
    operador_combobox = ttk.Combobox(janela_principal, values=["maior que", "menor que", "igual a"])
    operador_combobox.pack()

    tk.Label(janela_principal, text="Valor:").pack()
    valor_entry = tk.Entry(janela_principal)
    valor_entry.pack()

    botao_gerar_relatorio = tk.Button(janela_principal, text="Gerar Relatório", command=gerar_relatorio_callback)
    botao_gerar_relatorio.pack()

    botao_nova_busca = tk.Button(janela_principal, text="Nova Busca", command=nova_busca)
    botao_nova_busca.pack()

    janela_principal.mainloop()
