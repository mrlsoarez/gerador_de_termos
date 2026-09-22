
import tkinter as tk
from tkinter import ttk

from Env.env import inicializarAmbiente
from Modules.Planilha import analisarPlanilha, resetarPlanilha

from Classes.Documento import Termo, Relatorio

# Classes em português e camelCase 


def MAIN():

    INICIALIZADOR = inicializarAmbiente()  

    def GERAR_TERMO(dados, op):

        INICIALIZADOR.setModelo(op["opcao"])
        
        for chave in dados: 
            for index in range(len(dados[chave])): 
                termo = Termo(dados[chave][index], INICIALIZADOR.modelo, INICIALIZADOR.pastaAtual)
                termo.setOrdem(index+1)                
                termo.setEndereco(rf"{INICIALIZADOR.pastaAtual}\{dados[chave][index].tipo}")            
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
        
        relatorio = Relatorio(ARRAY, INICIALIZADOR.numeroProtocolo, INICIALIZADOR.modelo, INICIALIZADOR.pastaAtual)
        relatorio.setEndereco(rf"{relatorio.endereco.rsplit("\\", 1)[0]}\PROTOCOLOS")
        relatorio.criarArquivo(ARRAY)
        pass 
    
    def REINICIAR_PLANILHA(dados, op): 
        marcados = VISUAL_PLANILHA(dados)
        resetarPlanilha(marcados, INICIALIZADOR)
        INICIALIZADOR.setPlanilha()
        dados = analisarPlanilha(INICIALIZADOR)
        INICIALIZADOR.copiarTermosAnteriores(dados)
        
    gerarInformacao = {
        1: GERAR_TERMO,
        2: GERAR_RELATORIO,
        4: REINICIAR_PLANILHA
    }
    
    while True: 
        
        op = VISUAL(INICIALIZADOR)

        if (op["opcao"] == 1 or op["opcao"] == 2 or op["opcao"] == 4): 
            dados = analisarPlanilha(INICIALIZADOR) 
            gerarInformacao[op["opcao"]](dados, op) 
            break
        elif (op["opcao"] == 3):
            INICIALIZADOR.incrementarProtocolo()
            INICIALIZADOR = inicializarAmbiente()
        elif (op["opcao"] == 5): 
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
        text=f"N° PROTOCOLO: {INICIALIZADOR.numeroProtocolo}",
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
        "4. REINICIAR PLANILHA",
        "5. ENCERRAR"
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

def VISUAL_PLANILHA(dados): 
    LARGURA = 650
    ALTURA = 500

    FONTE_TITULO = ("Segoe UI", 19, "bold")
    FONTE_LABEL = ("Segoe UI", 11, "bold")
    FONTE_OPCAO = ("Segoe UI", 10)

    ROOT = tk.Tk()

    ROOT.title("Seleção de Itens")
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
        text="SELEÇÃO DE ITENS",
        font=FONTE_TITULO
    ).pack(side="left")


    ttk.Separator(
        ROOT,
        orient="horizontal"
    ).pack(
        fill="x",
        padx=35,
        pady=(5, 20)
    )


    # ============================================================
    # INSTRUÇÃO
    # ============================================================

    tk.Label(
        ROOT,
        text="Marque os itens que deseja processar",
        font=FONTE_LABEL,
        anchor="w"
    ).pack(
        fill="x",
        padx=55,
        pady=(0, 10)
    )


    # ============================================================
    # ÁREA COM SCROLL
    # ============================================================

    container = tk.Frame(ROOT)
    container.pack(
        fill="both",
        expand=True,
        padx=45,
        pady=(0, 10)
    )

    # Canvas
    canvas = tk.Canvas(
        container,
        highlightthickness=0
    )

    # Scrollbar
    scrollbar = ttk.Scrollbar(
        container,
        orient="vertical",
        command=canvas.yview
    )

    # Frame que ficará dentro do Canvas
    frame_itens = tk.Frame(canvas)

    frame_itens.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window(
        (0, 0),
        window=frame_itens,
        anchor="nw"
    )

    canvas.configure(
        yscrollcommand=scrollbar.set
    )

    canvas.pack(
        side="left",
        fill="both",
        expand=True
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )


    # ============================================================
    # CHECKBOXES
    # ============================================================

    checkboxes = []

    for chave in dados:

        # Título da categoria
        tk.Label(
            frame_itens,
            text=chave.upper(),
            font=FONTE_LABEL,
            anchor="w"
        ).pack(
            fill="x",
            pady=(8, 5)
        )

        for dado in dados[chave]:

            marcado = tk.BooleanVar(value=False)

            texto = f"{dado.contratado} - {dado.af}"

            checkbox = tk.Checkbutton(
                frame_itens,
                text=texto,
                variable=marcado,
                font=FONTE_OPCAO,
                anchor="w",
                padx=5,
                pady=3
            )

            checkbox.pack(
                fill="x"
            )

            checkboxes.append(
                (chave, dado, marcado)
            )


    # ============================================================
    # RESULTADO
    # ============================================================

    resultado = {}


    def confirmar():

        resultado.clear()

        for chave in dados:

            resultado[chave] = []

        for chave, dado, marcado in checkboxes:

            # Marcado = TRUE
            # Não entra no array

            if marcado.get() is False:

                # Desmarcado = FALSE
                # Entra no array

                resultado[chave].append(dado.ordem)

        ROOT.destroy()


    # ============================================================
    # BOTÃO
    # ============================================================

    ttk.Separator(
        ROOT,
        orient="horizontal"
    ).pack(
        fill="x",
        padx=35,
        pady=(5, 10)
    )

    tk.Button(
        ROOT,
        text="CONFIRMAR",
        width=18,
        height=2,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        command=confirmar
    ).pack(
        pady=(0, 15)
    )


    # ============================================================
    # MOUSE WHEEL
    # ============================================================

    def scroll(event):

        canvas.yview_scroll(
            int(-1 * (event.delta / 120)),
            "units"
        )

    canvas.bind_all(
        "<MouseWheel>",
        scroll
    )


    ROOT.mainloop()

    return resultado

MAIN()

