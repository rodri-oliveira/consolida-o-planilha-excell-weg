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

    # Criação da janela principal
    janela_principal = tk.Tk()
    janela_principal.title("Consolidar Planilhas do SharePoint")

    tipo_selecao = tk.StringVar(value="TIN")  # Valor padrão para o botão de rádio
    opcoes_consolidacao = {
        "Alocação": tk.BooleanVar(value=False),
        "Backlog": tk.BooleanVar(value=False),
        "Horas Disponíveis": tk.BooleanVar(value=False),
    }

    # Grupo de Botões de Rádio para seleção de tipo
    frame_radio = tk.Frame(janela_principal)
    frame_radio.pack(pady=10)
    tk.Label(frame_radio, text="Selecione o tipo:").pack(side=tk.LEFT)
    tk.Radiobutton(frame_radio, text="SEG", variable=tipo_selecao, value="SEG").pack(side=tk.LEFT)
    tk.Radiobutton(frame_radio, text="SGI", variable=tipo_selecao, value="SGI").pack(side=tk.LEFT)
    tk.Radiobutton(frame_radio, text="TIN", variable=tipo_selecao, value="TIN").pack(side=tk.LEFT)

    # Checkboxes para opções de consolidação
    frame_checkboxes = tk.Frame(janela_principal)
    frame_checkboxes.pack(pady=10)
    for opcao in opcoes_consolidacao:
        tk.Checkbutton(frame_checkboxes, text=opcao, variable=opcoes_consolidacao[opcao]).pack(anchor=tk.W)

    # Função para selecionar arquivos
    def selecionar_arquivos():
        listas = buscar_listas_sharepoint(token)
        if not listas or 'd' not in listas or 'results' not in listas['d']:
            messagebox.showerror("Erro", "Não foi possível buscar os arquivos do SharePoint.")
            return

        arquivos_disponiveis = listas['d']['results']

        # Filtrar arquivos de acordo com a seleção do botão de rádio
        arquivos_selecionados.clear()  # Limpa a lista anterior
        if tipo_selecao.get() == "TIN":
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

    # Função para consolidar
    def consolidar():
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
                messagebox.showinfo("Sucesso", "Consolidação das horas (Horas Backlog) realizada com sucesso!")
        else:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")

    # Botão para realizar a seleção de arquivos
    botao_selecionar_arquivos = tk.Button(janela_principal, text="Selecionar Arquivos", command=selecionar_arquivos)
    botao_selecionar_arquivos.pack(pady=10)

    # Botão para realizar a consolidação
    botao_consolidar = tk.Button(janela_principal, text="Consolidar", command=consolidar)
    botao_consolidar.pack(pady=10)

    # Botão para nova pesquisa
    botao_nova_pesquisa = tk.Button(janela_principal, text="Nova Pesquisa", command=lambda: [arquivos_selecionados.clear(), opcoes_consolidacao["Alocação"].set(False), opcoes_consolidacao["Backlog"].set(False), opcoes_consolidacao["Horas Disponíveis"].set(False)])
    botao_nova_pesquisa.pack()

    janela_principal.mainloop()
