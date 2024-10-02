import tkinter as tk
from tkinter import messagebox
from auth import obter_token_sharepoint, buscar_listas_sharepoint
from consolidar import (
    consolidar_planilhas_sharepoint,
    consolidar_aba_backlog_sharepoint,
    consolidar_horas_backlog_sharepoint,
)

def consolidar_planilhas_interface():
    arquivos_selecionados = []

    def selecionar_arquivos():
        token = obter_token_sharepoint()
        if not token:
            messagebox.showerror("Erro", "Não foi possível obter o token de acesso ao SharePoint.")
            return

        # Obtenha a lista de arquivos do SharePoint
        listas = buscar_listas_sharepoint(token)
        if not listas or 'd' not in listas or 'results' not in listas['d']:
            messagebox.showerror("Erro", "Não foi possível buscar os arquivos do SharePoint.")
            return

        arquivos_disponiveis = listas['d']['results']

        # Exibir lista para o usuário escolher os arquivos
        arquivos_selecionados.clear()  # Limpa a lista anterior
        for arquivo in arquivos_disponiveis:
            nome_arquivo = arquivo['Name']
            if nome_arquivo.startswith("TIN "):  # Verifique se o nome do arquivo começa com "TIN "
                incluir = messagebox.askyesno("Seleção de Arquivos", f"Incluir o arquivo {nome_arquivo} na consolidação?")
                if incluir:
                    arquivos_selecionados.append(arquivo)

        if arquivos_selecionados:
            messagebox.showinfo("Seleção de Arquivos", "Arquivos selecionados com sucesso!")
        else:
            messagebox.showinfo("Seleção de Arquivos", "Nenhum arquivo selecionado.")

    def consolidar_abas():
        print("Arquivos selecionados para consolidação:", arquivos_selecionados)  # Adicionando depuração
        if arquivos_selecionados:
            caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
            consolidar_planilhas_sharepoint(caminho_das_planilhas)
            messagebox.showinfo("Sucesso", "Consolidação das abas (exceto Backlog) realizada com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")

    def consolidar_backlog():
        print("Arquivos selecionados para consolidação de Backlog:", arquivos_selecionados)  # Adicionando depuração
        if arquivos_selecionados:
            caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
            consolidar_aba_backlog_sharepoint(caminho_das_planilhas)
            messagebox.showinfo("Sucesso", "Consolidação das abas Backlog realizada com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")

    def consolidar_horas_backlog():
        print("Arquivos selecionados para consolidação de Horas Backlog:", arquivos_selecionados)  # Adicionando depuração
        if arquivos_selecionados:
            caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
            consolidar_horas_backlog_sharepoint(caminho_das_planilhas)
            messagebox.showinfo("Sucesso", "Consolidação das horas Backlog realizada com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")

    def nova_pesquisa():
        arquivos_selecionados.clear()  # Limpa a lista
        messagebox.showinfo("Nova Pesquisa", "Seleção de arquivos reiniciada.")

    # Criação da janela principal
    janela_principal = tk.Tk()
    janela_principal.title("Consolidar Planilhas do SharePoint")

    # Botão para selecionar arquivos
    botao_selecionar_arquivos = tk.Button(janela_principal, text="Selecionar Arquivos", command=selecionar_arquivos)
    botao_selecionar_arquivos.pack()

    # Botão para consolidar todas as abas (menos a Backlog)
    botao_consolidar_abas = tk.Button(janela_principal, text="Consolidar Abas", command=consolidar_abas)
    botao_consolidar_abas.pack()

    # Botão para consolidar apenas Backlog
    botao_consolidar_backlog = tk.Button(janela_principal, text="Consolidar Backlog", command=consolidar_backlog)
    botao_consolidar_backlog.pack()

    # Botão para consolidar horas Backlog
    botao_consolidar_horas_backlog = tk.Button(janela_principal, text="Consolidar Horas Backlog", command=consolidar_horas_backlog)
    botao_consolidar_horas_backlog.pack()

    # Botão para nova pesquisa
    botao_nova_pesquisa = tk.Button(janela_principal, text="Nova Pesquisa", command=nova_pesquisa)
    botao_nova_pesquisa.pack()

    janela_principal.mainloop()

