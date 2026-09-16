import tkinter as tk
from tkinter import ttk

from Env.env import inicializarAmbiente
from Modules.Planilha import analisarPlanilha

from Classes.Documento import Termo

# Classes em português e camelCase 


def MAIN():

    INICIALIZADOR = inicializarAmbiente()  

    def GERAR_TERMO(dados, op):

        INICIALIZADOR.setModelo(op)

        for dado in dados:
            termo = Termo(dado, INICIALIZADOR.modelo, INICIALIZADOR.pasta_atual)
            print(termo)
        pass 

    def GERAR_RELATORIO(dados, op):
        INICIALIZADOR.setModelo(op)
        pass 

    gerarInformacao = {
        1: GERAR_TERMO,
        2: GERAR_RELATORIO
    }
    
    while True: 
        #op = VISUAL(INICIALIZADOR)
        op = 1
        if (op == 1 or op == 2): 
            dados = analisarPlanilha(INICIALIZADOR) 
            gerarInformacao[op](dados, op) 
            break
        if (op == 4): 
            break 
    
def VISUAL(INICIALIZADOR):

    # ============================================================
    # CONFIGURAÇÕES
    # ============================================================

    LARGURA = 650
    ALTURA = 450

    FONTE_TITULO = ("Segoe UI", 19, "bold")
    FONTE_PROTOCOLO = ("Segoe UI", 10, "bold")
    FONTE_LABEL = ("Segoe UI", 11, "bold")
    FONTE_OPCAO = ("Segoe UI", 11)
    FONTE_BOTAO = ("Segoe UI", 10, "bold")
    FONTE_RODAPE = ("Segoe UI", 8)

    # ============================================================
    # JANELA
    # ============================================================

    ROOT = tk.Tk()

    ROOT.title("Gerador de Termos")
    ROOT.geometry(f"{LARGURA}x{ALTURA}")
    ROOT.resizable(False, False)

    # ============================================================
    # HEADER
    # ============================================================

    header = tk.Frame(ROOT)
    header.pack(
        fill="x",
        padx=35,
        pady=(25, 10)
    )

    tk.Label(
        header,
        text="GERADOR DE TERMOS",
        font=FONTE_TITULO
    ).pack(side="left")

    tk.Label(
        header,
        text=f"N° PROTOCOLO: {INICIALIZADOR.numero_protocolo}",
        font=FONTE_PROTOCOLO
    ).pack(side="right")

    # Linha
    ttk.Separator(
        ROOT,
        orient="horizontal"
    ).pack(
        fill="x",
        padx=35,
        pady=(5, 25)
    )

    # ============================================================
    # OPÇÕES
    # ============================================================

    conteudo = tk.Frame(ROOT)
    conteudo.pack(
        fill="x",
        padx=55
    )

    tk.Label(
        conteudo,
        text="Selecione uma opção",
        font=FONTE_LABEL,
        anchor="w"
    ).pack(
        fill="x",
        pady=(0, 12)
    )

    opcao = tk.IntVar(value=0)

    opcoes = [
        "1. GERAR TERMOS",
        "2. GERAR RELATÓRIO",
        "3. ATUALIZAR PROTOCOLO",
        "4. ENCERRAR"
    ]

    for i, texto in enumerate(opcoes):

        tk.Radiobutton(
            conteudo,
            text=texto,
            variable=opcao,
            value=i + 1,
            font=FONTE_OPCAO,
            anchor="w",
            padx=10,
            pady=6,
            cursor="hand2"
        ).pack(
            fill="x"
        )

    # ============================================================
    # BOTÃO
    # ============================================================

    def enviar():
        ROOT.destroy()

    ttk.Separator(
        ROOT,
        orient="horizontal"
    ).pack(
        fill="x",
        padx=35,
        pady=(25, 15)
    )

    tk.Button(
        ROOT,
        text="ENVIAR",
        width=18,
        height=2,
        font=FONTE_BOTAO,
        cursor="hand2",
        command=enviar
    ).pack()

    # ============================================================
    # EXECUÇÃO
    # ============================================================

    ROOT.mainloop()

    return opcao.get()

MAIN()

