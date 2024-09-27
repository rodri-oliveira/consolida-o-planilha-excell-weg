import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from consolidar import consolidar_planilhas  # Certifique-se de que essa linha esteja correta

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

    janela_principal.mainloop()