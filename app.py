import tkinter as tk
from tkinter import messagebox
from auth import obter_token_sharepoint, buscar_listas_sharepoint
from consolidar import (
    consolidar_planilhas_sharepoint,
    consolidar_aba_backlog_sharepoint,
    consolidar_horas_backlog_sharepoint,
)

token = obter_token_sharepoint()

def consolidar_planilhas_interface():
    arquivos_selecionados = []

    # Função para selecionar arquivos com base no prefixo
    def selecionar_arquivos_por_prefixo(prefixo):
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
            messagebox.showinfo("Seleção de Arquivos", f"Arquivos que começam com {prefixo} selecionados com sucesso!")
        else:
            messagebox.showinfo("Seleção de Arquivos", "Nenhum arquivo selecionado.")

    def consolidar_abas():
        print("consolidar_abas() no app.py")
        if arquivos_selecionados:
            if opcoes_consolidacao["Alocação"].get():
                caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
                consolidar_planilhas_sharepoint(caminho_das_planilhas, token)
                messagebox.showinfo("Sucesso", "Consolidação das abas (exceto Backlog) realizada com sucesso!")

            if opcoes_consolidacao["Backlog"].get():
                caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
                consolidar_aba_backlog_sharepoint(caminho_das_planilhas, token)
                messagebox.showinfo("Sucesso", "Consolidação das abas Backlog realizada com sucesso!")

            if opcoes_consolidacao["Horas Disponíveis"].get():
                caminho_das_planilhas = [arquivo['ServerRelativeUrl'] for arquivo in arquivos_selecionados]
                consolidar_horas_backlog_sharepoint(caminho_das_planilhas, token)
                messagebox.showinfo("Sucesso", "Consolidação das horas Backlog realizada com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")

    # Botão para realizar a consolidação
    botao_consolidar = tk.Button(janela_principal, text="Consolidar", command=consolidar)
    botao_consolidar.pack(pady=10)

    # Botão para nova pesquisa
    botao_nova_pesquisa = tk.Button(janela_principal, text="Nova Pesquisa", command=lambda: [arquivos_selecionados.clear(), opcoes_consolidacao["Alocação"].set(False), opcoes_consolidacao["Backlog"].set(False), opcoes_consolidacao["Horas Disponíveis"].set(False)])
    botao_nova_pesquisa.pack()

    janela_principal.mainloop()

