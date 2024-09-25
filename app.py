# Importa as bibliotecas necessárias
import os  # Para manipulação de caminhos e diretórios
import pandas as pd  # Para manipulação de dados em DataFrames
import tkinter as tk  # Para a criação da interface gráfica
from tkinter import messagebox, ttk, filedialog  # Para mensagens e widgets da interface gráfica

# Variável global para armazenar o DataFrame consolidado
dataframe_consolidado = None

def consolidar_planilhas(caminho_pasta_origem, caminho_saida):
    global dataframe_consolidado  # Acessa a variável global
    lista_dataframes = []  # Lista para armazenar os DataFrames das planilhas

    try:
        # Percorre todos os arquivos na pasta de origem
        for nome_arquivo in os.listdir(caminho_pasta_origem):
            if nome_arquivo.endswith('.xlsx'):  # Verifica se o arquivo é uma planilha Excel
                caminho_arquivo_excel = os.path.join(caminho_pasta_origem, nome_arquivo)  # Cria o caminho completo do arquivo
                xls = pd.ExcelFile(caminho_arquivo_excel)  # Lê o arquivo Excel

                # Percorre todas as abas do arquivo Excel
                for nome_aba in xls.sheet_names:
                    # Lê a aba e armazena em um DataFrame
                    dataframe_aba = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_aba)
                    lista_dataframes.append(dataframe_aba)  # Adiciona o DataFrame à lista

        # Verifica se a lista de DataFrames não está vazia
        if lista_dataframes:
            # Concatena todos os DataFrames em um único DataFrame
            dataframe_consolidado = pd.concat(lista_dataframes, ignore_index=True)

            # Define o caminho para salvar o DataFrame consolidado
            caminho_saida = "C:/consolida-o-planilha-excell-weg/planilhas-consolidadas/planilha_consolidada.xlsx"
            # Cria os diretórios necessários, se não existirem
            os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
            # Salva o DataFrame consolidado em um arquivo Excel
            dataframe_consolidado.to_excel(caminho_saida, index=False)
            print(f'Planilhas consolidadas com sucesso! Salvas em {caminho_saida}')
        else:
            print("Nenhuma planilha foi encontrada ou consolidada.")  # Mensagem se não houver planilhas

    except FileNotFoundError:
        # Exibe uma mensagem de aviso se o caminho de origem não for encontrado
        messagebox.showwarning("Erro", f"O caminho '{caminho_pasta_origem}' não foi encontrado.")
    except pd.errors.EmptyDataError:
        # Mensagem de aviso se um dos arquivos Excel estiver vazio
        messagebox.showwarning("Erro", "Um dos arquivos Excel está vazio.")
    except Exception as e:
        # Captura e exibe qualquer outro erro que ocorra
        messagebox.showwarning("Erro", f"Ocorreu um erro: {str(e)}")

def gerar_relatorio(dataframe, nome_coluna, operador, valor):
    try:
        # Verifica se a coluna existe no DataFrame
        if nome_coluna not in dataframe.columns:
            messagebox.showwarning("Erro", f"A coluna '{nome_coluna}' não existe no DataFrame.")
            return pd.DataFrame()  # Retorna um DataFrame vazio

        # Converte a coluna para numérico, tratando erros
        dataframe[nome_coluna] = pd.to_numeric(dataframe[nome_coluna], errors='coerce')

        # Aplica o filtro de acordo com o operador selecionado
        if operador == "maior que":
            return dataframe[dataframe[nome_coluna] > valor]
        elif operador == "menor que":
            return dataframe[dataframe[nome_coluna] < valor]
        elif operador == "igual a":
            return dataframe[dataframe[nome_coluna] == valor]
        return pd.DataFrame()  # Retorna um DataFrame vazio se nenhum operador corresponder
    except KeyError:
        # Mensagem se a coluna não for encontrada
        messagebox.showwarning("Erro", f"A coluna '{nome_coluna}' não foi encontrada.")
        return pd.DataFrame()  # Retorna um DataFrame vazio
    except Exception as e:
        # Captura e exibe qualquer erro que ocorra
        messagebox.showwarning("Erro", f"Ocorreu um erro ao aplicar o filtro: {str(e)}")
        return pd.DataFrame()  # Retorna um DataFrame vazio

def consolidar_planilhas_interface():
    # Função para consolidar planilhas quando o botão é pressionado
    def consolidar():
        caminho_pasta_origem = entrada_pasta.get()  # Obtém o caminho da pasta do campo de entrada
        consolidar_planilhas(caminho_pasta_origem, None)  # Chama a função de consolidação

        # Atualiza as opções do combobox de colunas se o DataFrame consolidado não for None
        if dataframe_consolidado is not None:
            coluna_combobox['values'] = dataframe_consolidado.columns.tolist()

        # Mensagem de sucesso
        messagebox.showinfo("Sucesso", "Planilhas consolidadas com sucesso!")

    # Função de callback para gerar o relatório
    def gerar_relatorio_callback():
        nome_coluna = coluna_combobox.get()  # Obtém o nome da coluna selecionada
        operador = operador_combobox.get()  # Obtém o operador selecionado
        valor = valor_entry.get()  # Obtém o valor de entrada

        # Verifica se todos os campos estão preenchidos
        if nome_coluna and operador and valor:
            try:
                valor = float(valor)  # Converte o valor para float
                relatorio_filtrado = gerar_relatorio(dataframe_consolidado, nome_coluna, operador, valor)  # Chama a função de gerar relatório

                # Verifica se o DataFrame filtrado está vazio
                if relatorio_filtrado.empty:
                    messagebox.showinfo("Resultado", "Nenhum dado encontrado para os critérios selecionados.")
                else:
                    # Define o caminho para salvar o relatório gerado
                    caminho_saida_relatorio = "C:/consolida-o-planilha-excell-weg/relatórios/relatorio.xlsx"
                    # Cria os diretórios necessários, se não existirem
                    os.makedirs(os.path.dirname(caminho_saida_relatorio), exist_ok=True)

                    try:
                        # Salva o DataFrame filtrado em um arquivo Excel
                        relatorio_filtrado.to_excel(caminho_saida_relatorio, index=False)
                        messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em {caminho_saida_relatorio}!")
                    except FileNotFoundError:
                        # Mensagem de aviso se o caminho para salvar o relatório não for encontrado
                        messagebox.showwarning("Erro", "O caminho para salvar o relatório não foi encontrado.")
                    except Exception as e:
                        # Mensagem de erro ao salvar o relatório
                        messagebox.showwarning("Erro", f"Ocorreu um erro ao salvar o relatório: {str(e)}")
            except ValueError:
                # Mensagem de erro se o valor não for numérico
                messagebox.showwarning("Erro", "Insira um valor numérico válido.")
        else:
            # Mensagem de aviso se algum campo não estiver preenchido
            messagebox.showwarning("Atenção", "Preencha todos os campos antes de gerar o relatório.")

    # Função para buscar o caminho da pasta
    def buscar_pasta():
        caminho_selecionado = filedialog.askdirectory()  # Abre um diálogo para selecionar um diretório
        if caminho_selecionado:
            entrada_pasta.delete(0, tk.END)  # Limpa o campo de entrada
            entrada_pasta.insert(0, caminho_selecionado)  # Insere o caminho selecionado no campo de entrada

    # Função para reiniciar a busca
    def nova_busca():
        entrada_pasta.delete(0, tk.END)  # Limpa o campo de entrada
        coluna_combobox.set('')  # Limpa o combobox de colunas
        operador_combobox.set('')  # Limpa o combobox de operadores
        valor_entry.delete(0, tk.END)  # Limpa o campo de entrada do valor

    # Criação da janela principal da interface gráfica
    janela_principal = tk.Tk()
    janela_principal.title("Consolidar Planilhas")  # Define o título da janela

    # Label e campo de entrada para o caminho da pasta
    tk.Label(janela_principal, text="Caminho da pasta:").pack()  # Label para o caminho da pasta
    entrada_pasta = tk.Entry(janela_principal, width=50)  # Campo de entrada para o caminho
    entrada_pasta.pack()  # Adiciona o campo à janela

    # Botão para buscar a pasta
    botao_buscar_pasta = tk.Button(janela_principal, text="Buscar Pasta", command=buscar_pasta)
    botao_buscar_pasta.pack()  # Adiciona o botão à janela

    # Botão para consolidar planilhas
    botao_consolidar = tk.Button(janela_principal, text="Consolidar", command=consolidar)
    botao_consolidar.pack()  # Adiciona o botão à janela

    # Label e combobox para seleção de colunas
    tk.Label(janela_principal, text="Selecione uma coluna:").pack()
    coluna_combobox = ttk.Combobox(janela_principal)  # Combobox para seleção de colunas
    coluna_combobox.pack()  # Adiciona o combobox à janela

    # Label e combobox para seleção de operadores
    tk.Label(janela_principal, text="Selecione um operador:").pack()
    operador_combobox = ttk.Combobox(janela_principal, values=["maior que", "menor que", "igual a"])  # Combobox para operadores
    operador_combobox.pack()  # Adiciona o combobox à janela

    # Label e campo de entrada para o valor
    tk.Label(janela_principal, text="Valor:").pack()
    valor_entry = tk.Entry(janela_principal)  # Campo de entrada para o valor
    valor_entry.pack()  # Adiciona o campo à janela

    # Botão para gerar o relatório
    botao_gerar_relatorio = tk.Button(janela_principal, text="Gerar Relatório", command=gerar_relatorio_callback)
    botao_gerar_relatorio.pack()  # Adiciona o botão à janela

    # Botão para nova busca
    botao_nova_busca = tk.Button(janela_principal, text="Nova Busca", command=nova_busca)
    botao_nova_busca.pack()  # Adiciona o botão à janela

    # Inicia a interface gráfica
    janela_principal.mainloop()


