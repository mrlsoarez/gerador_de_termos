import tkinter as tk
from tkinter import ttk


class GeradorTermosGUI:
    def __init__(self, root):
        
       

     
        # Definindo a janela principal
        self.root = root
        self.root.title("APP")
        self.root.geometry("450x250")
        self.root.resizable(False, False)

        # Variáveis 
        self.tipo = tk.StringVar(value="1")
        self.atualizar = tk.BooleanVar(value=False)
        self.geracao_total = tk.BooleanVar(value=False)
        #resposta = CAPTURAR_RESPOSTA("Bem vindo! Escolha dentre as opções \n1. Gerar termos aditivos\n2. Gerar relatório\n3. Gerar portaria\n4. Atualizar número de protocolo\n5. Encerrar\n-> ", ("1", "2", "3", "4", "5"))
        
        ttk.Label(
            root,
            text="Bem-vindo!",
            font=("Segoe UI", 14, "bold"),
        ).pack(pady=10)
        
        ttk.Label(
            root,
            text="Escolha o que deseja fazer:",
            font=("Segoe UI", 10),
        ).pack(pady=10)
        
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill="both", expand=True)
        
        ttk.Radiobutton(
                frame,
                text="1. Gerar Termos Aditivos",
                variable=self.tipo,
                value="1",
        ).pack(anchor="w")
        
        ttk.Radiobutton(
            frame,
            text="2. Gerar Relatório",
            variable=self.tipo,
            value="2",
        ).pack(anchor="w")
        
        ttk.Radiobutton(
            frame,
            text="3. Gerar Portaria",
            variable=self.tipo,
            value="3",
        ).pack(anchor="w")
        
        ttk.Radiobutton(
            frame,
            text="4. Atualizar Número de Protocolo",
            variable=self.tipo,
            value="4",
        ).pack(anchor="w")
        
        ttk.Radiobutton(
            frame,
            text="5. Encerrar",
            variable=self.tipo,
            value="5",
        ).pack(anchor="w")
        
        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=10)
        
        
        ttk.Button(
            frame,
            text="Gerar"
        ).pack(pady=20)
        
    def confirmar(self):
        self.pergunta_inicial = self.tipo.get()                  # "1" ou "2"
        self.pergunta_update = "s" if self.atualizar.get() else "n"
        self.pergunta_total = "s" if self.geracao_total.get() else "n"
        """
        
     

        # Pergunta 1
        ttk.Label(frame, text="Escolha o tipo de documento:").pack(anchor="w")

        ttk.Radiobutton(
            frame,
            text="Contrato",
            variable=self.tipo,
            value="1",
        ).pack(anchor="w")

        ttk.Radiobutton(
            frame,
            text="Ata",
            variable=self.tipo,
            value="2",
        ).pack(anchor="w")


        # Pergunta 2
        ttk.Checkbutton(
            frame,
            text="Atualizar os dados da planilha",
            variable=self.atualizar,
        ).pack(anchor="w")

        # Pergunta 3
        ttk.Checkbutton(
            frame,
            text="Geração total",
            variable=self.geracao_total,
        ).pack(anchor="w")





        """
root = tk.Tk()
app = GeradorTermosGUI(root)
root.mainloop()
"""

# Valores equivalentes ao seu CAPTURAR_RESPOSTA
pergunta_inicial = app.pergunta_inicial
pergunta_update = app.pergunta_update
pergunta_total = app.pergunta_total

print(pergunta_inicial)
print(pergunta_update)
print(pergunta_total)
"""