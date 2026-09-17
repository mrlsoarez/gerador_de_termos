
import tkinter as tk
from tkinter import ttk

from Env.env import inicializarAmbiente
from Modules.Planilha import analisarPlanilha

from Classes.Documento import Termo, Relatorio

# Classes em português e camelCase 


def MAIN():

    INICIALIZADOR = inicializarAmbiente()  

    def GERAR_TERMO(dados, op):

        INICIALIZADOR.setModelo(op["opcao"])
        
        for chave in dados: 
            for index in range(len(dados[chave])): 
                termo = Termo(dados[chave][index], INICIALIZADOR.modelo, INICIALIZADOR.pasta_atual)
                termo.setOrdem(index+1)                
                termo.setEndereco(rf"{INICIALIZADOR.pasta_atual}\{dados[chave][index].tipo}")            
                if not termo.verificarSeExiste():
                    termo.criarArquivo()

    def GERAR_RELATORIO(dados, op):
        
        INICIALIZADOR.setModelo(op["opcao"])
        ARRAY = []
        
        if (op["tipo_relatorio"] == 1):
            ARRAY = dados['contrato']
        if (op["tipo_relatorio"] == 2):
            ARRAY = dados['ata']
        if (op["tipo_relatorio"] == 3):
            ARRAY = dados['contrato']
            for dado in dados['ata']:
                ARRAY.append(dado)
        
        relatorio = Relatorio(ARRAY, INICIALIZADOR.numero_protocolo, INICIALIZADOR.modelo, INICIALIZADOR.pasta_atual)
        relatorio.setEndereco(rf"{relatorio.endereco.rsplit("\\", 1)[0]}\PROTOCOLOS")
        relatorio.criarArquivo(ARRAY)
        pass 

    gerarInformacao = {
        1: GERAR_TERMO,
        2: GERAR_RELATORIO
    }
    
    while True: 
        op = VISUAL(INICIALIZADOR)
        if (op["opcao"] == 1 or op["opcao"] == 2): 
            dados = analisarPlanilha(INICIALIZADOR) 
            gerarInformacao[op["opcao"]](dados, op) 
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
    # SEGUNDA TELA - TIPO DE RELATÓRIO
    # ============================================================

    def segunda_tela():

        janela = tk.Toplevel(ROOT)

        janela.title("Tipo de Relatório")
        janela.geometry("500x350")
        janela.resizable(False, False)

        # Impede interação com a primeira tela
        janela.grab_set()

        tk.Label(
            janela,
            text="TIPO DE RELATÓRIO",
            font=FONTE_TITULO
        ).pack(
            pady=(30, 20)
        )

        tipo_relatorio = tk.IntVar(value=0)

        opcoes_relatorio = [
            "1. CONTRATO",
            "2. ATA",
            "3. AMBOS"
        ]

        for i, texto in enumerate(opcoes_relatorio):

            tk.Radiobutton(
                janela,
                text=texto,
                variable=tipo_relatorio,
                value=i + 1,
                font=FONTE_OPCAO,
                anchor="w",
                padx=10,
                pady=6,
                cursor="hand2"
            ).pack(
                fill="x",
                padx=55
            )

        def confirmar():

            if tipo_relatorio.get() == 0:
                return

            resultado["tipo_relatorio"] = tipo_relatorio.get()

            janela.destroy()
            ROOT.destroy()

        tk.Button(
            janela,
            text="CONFIRMAR",
            width=18,
            height=2,
            font=FONTE_BOTAO,
            cursor="hand2",
            command=confirmar
        ).pack(
            pady=25
        )


    # ============================================================
    # BOTÃO
    # ============================================================

    resultado = {
        "opcao": None,
        "tipo_relatorio": None
    }


    def enviar():

        if opcao.get() == 0:
            return

        resultado["opcao"] = opcao.get()

        # Se for GERAR RELATÓRIO
        if opcao.get() == 2:
            segunda_tela()

        else:
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

    return resultado

MAIN()

