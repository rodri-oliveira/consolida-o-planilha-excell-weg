import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from consolidar import consolidar_planilhas

class ConsolidarPlanilhasApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Consolidar Planilhas")
        self.btn_selecionar_pasta = tk.Button(master, text="Selecionar Pasta", command=self.selecionar_pasta)
        self.btn_selecionar_pasta.pack(pady=10)
        self.label_pasta = tk.Label(master, text="")
        self.label_pasta.pack(pady=10)
        self.btn_consolidar = tk.Button(master, text="Consolidar Planilhas", command=self.consolidar)
        self.btn_consolidar.pack(pady=10)
        self.caminho_pasta_planilhas = ""

    def selecionar_pasta(self):
        self.caminho_pasta_planilhas = filedialog.askdirectory()
        self.label_pasta.config(text=self.caminho_pasta_planilhas)

    def consolidar(self):
        if not self.caminho_pasta_planilhas:
            messagebox.showerror("Erro", "Selecione uma pasta primeiro.")
            return
        try:
            consolidar_planilhas(self.caminho_pasta_planilhas)
            messagebox.showinfo("Sucesso", "Planilhas consolidadas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
