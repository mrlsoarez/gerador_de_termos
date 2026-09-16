# Importação de dados de ambiente, serviços externos e módulos reutilizáveis
from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, atualizar_numero_protocolo
from SERVICES.PLANILHA import ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS, INICIAR_PLANILHA
from SERVICES.DADOS_EXTERNOS import COLETAR_DADOS_EXTERNOS
from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

# Importação de Bibliotecas
import os 
import tkinter as tk
from tkinter import ttk

def MAIN():
    
    def iniciar_modulo_termos(op):
        
        def realizar_perguntas_iniciais():
            pergunta_inicial = CAPTURAR_RESPOSTA(
                "Bem vindo ao gerador de termos! Escolha dentre as opções para gerar: \n1. Contrato\n2. Ata\n-> ", ("1", "2")
            )
            pergunta_update = CAPTURAR_RESPOSTA(
                "Deseja atualizar os dados da planilha? (S/N) -> ", ("s", "n")
            )
            pergunta_total = CAPTURAR_RESPOSTA(
                "Deseja uma geração total? (S/N) -> ", ("s", "n")
            )
            return pergunta_inicial, pergunta_update, pergunta_total
                
        def criar_pasta_termos(gerenciador):
            
            mes_numero = EncontrarData("mes", False)
            mes_extenso = EncontrarData("mes", True) 
            dia = EncontrarData("dia")
        
            os.chdir(gerenciador.pasta_base)

            PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
            PASTA_DIA = f"{dia}-{mes_numero}"
            PROTOCOLO = f"REMESSA X - PROTOCOLO N° {gerenciador.numero_protocolo}"
        
            gerenciador.criarPasta(PASTA_MES, True)
            gerenciador.criarPasta(PASTA_DIA, True)
            gerenciador.criarPasta(PROTOCOLO, True)
            gerenciador.criarPasta("PROTOCOLOS", False)
            gerenciador.criarPasta("TERMOS", True)
            gerenciador.criarPasta(gerenciador.tipo_arquivo, True)
            gerenciador.criarPasta("WORD")
            gerenciador.criarPasta("PDF")
        
            gerenciador.pasta_atual = rf"{gerenciador.pasta_base}\{PASTA_MES}\{PASTA_DIA}\{PROTOCOLO}\TERMOS\{gerenciador.tipo_arquivo}"
           
        def capturar_dados_externos(option):
            print("Iniciando a coleta de dados externos.. por favor, aguarde..")
            try:
                if (option == 1): 
                    dados_contratos = COLETAR_DADOS_EXTERNOS("contratos")
                dados_licit = COLETAR_DADOS_EXTERNOS("licitacao")
                #dados_servidores =  COLETAR_DADOS_EXTERNOS("servidores")
                #dados_empenhos = COLETAR_DADOS_EXTERNOS("empenhos")
                #dados_liquidacao =  COLETAR_DADOS_EXTERNOS("liquidacoes")
                pass
            except: 
                print("Algo deu errado no processo de busca de dados")
            else:
                print("Dados encontrados e armazenados para inserção nas planilhas")

            print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
            print("Inserindo as informações na planilha de análise base.. favor, aguardar.")     
            try: 
                if (option == 1):
                    ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Base", dados_contratos)
                ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Licit", dados_licit)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Servidores", dados_servidores)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Empenhos", dados_empenhos)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Liquidacoes", dados_liquidacao)
            except:
                print("Algo deu errado no processo de inserir as informações nas planilhas")
            else: 
                print("Dados inseridos nas planilhas.")
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")

        pergunta_inicial = op["dados_extras"]["tipo_termo"]
        pergunta_update = op["dados_extras"]["atualizar_dados"]
        pergunta_total = op["dados_extras"]["analise_imediata"]
                    
        TIPO_TERMO = pegar_tipo_termo(pergunta_inicial)[0]
        PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO["arquivo"])
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), TIPO_TERMO["tipo"])
            
        criar_pasta_termos(GERENCIADOR_PASTAS)
        print(pergunta_update, pergunta_inicial, type(pergunta_inicial))
        if (pergunta_update == "s"): 
            capturar_dados_externos(pergunta_inicial)

        parametros = {
            "op": op["opcao"], 
            "modo_total": True if pergunta_total == "s" else False
        }
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■\nIniciando planilha..")
        INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)
    
    def iniciar_modulo_protocolo(op):
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
        
        def entrar_na_pasta_de_protocolo():
            mes_numero = EncontrarData("mes", False)
            mes_extenso = EncontrarData("mes", True) 
            dia = EncontrarData("dia")       
            
            PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
            PASTA_DIA = f"{dia}-{mes_numero}"
            
            return  GERENCIADOR_PASTAS.entrarEmPasta(fr"{PASTA_MES}\{PASTA_DIA}\REMESSA X - PROTOCOLO N° {GERENCIADOR_PASTAS.numero_protocolo}\PROTOCOLOS")
            
        def gerar_protocolo():
 
            ARRAY = []
            TIPO_TERMO = pegar_tipo_termo(op["dados_extras"]["tipo_relatorio"])
            
            parametros = {
                "op": op["opcao"], 
                "ARR": ARRAY
            }
            
            for i in range(len(TIPO_TERMO)):
                PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO[i]["arquivo"])
                GERENCIADOR_PASTAS.set_tipo_arquivo = TIPO_TERMO[i]["tipo"]
                ARRAY = INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)
            
            protocolo = ARRAY[0]
            
            protocolo.copiar_arquivo(protocolo.modelo, protocolo.endereco)
            protocolo.criar_arquivo(ARRAY)
            
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), "")
        
        if (not entrar_na_pasta_de_protocolo()):
            print("Pasta de relatório ainda não existe!")
            return
        
        gerar_protocolo()
        

        pass 
    
    def iniciar_modulo_portaria(op):
        parametros = {
            "op": op["opcao"]
        }
        TIPO_TERMO = pegar_tipo_termo(2)
        PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO[0]["arquivo"])
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(op), "")
        INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)
        pass 
    
    iniciador = VISUAL()
    opcao = iniciador["opcao"]
    
    while True: 
        if (opcao == 1):
            iniciar_modulo_termos(iniciador)
        elif (opcao == 2):
            iniciar_modulo_protocolo(iniciador)
        elif (opcao == 3):
            iniciar_modulo_portaria(iniciador)
        elif (opcao == 4):
            atualizar_numero_protocolo()
        elif (opcao == 5):
            print("Encerrando.")
            break
    
        iniciador = VISUAL()
        opcao = iniciador["opcao"]
    """
    
    while True:        
        resposta = CAPTURAR_RESPOSTA("Bem vindo! Escolha dentre as opções \n1. Gerar termos aditivos\n2. Gerar relatório\n3. Gerar portaria\n4. Atualizar número de protocolo\n5. Encerrar\n-> ", ("1", "2", "3", "4", "5"))
        opcao = resposta[0]
        if (opcao == "1"):
            iniciar_modulo_termos(opcao)
        elif (opcao == "2"):
            iniciar_modulo_protocolo(opcao)
        elif (opcao == "3"):
            iniciar_modulo_portaria(opcao)
        elif (opcao == "4"):
            atualizar_numero_protocolo()
        elif (opcao == "5"):
            print("Encerrando.")
            break
    """
    
def CAPTURAR_RESPOSTA(mensagem, dado_esperado):

    PERGUNTA = input(mensagem).lower()

    while PERGUNTA not in dado_esperado:
        print(("Opção não válida, por favor, digite uma das opções ao lado ", dado_esperado))
        PERGUNTA = input(mensagem).lower()
        
    return (PERGUNTA, True)

def VISUAL():
    
    dados_retorno = {
        "opcao": "",
        "dados_extras": ""
    }

    # ============================================================
    # CONFIGURAÇÕES VISUAIS
    # ============================================================

    LARGURA = 850
    ALTURA = 600

    FONTE_TITULO = ("Segoe UI", 20, "bold")
    FONTE_PROTOCOLO = ("Segoe UI", 11, "bold")
    FONTE_LABEL = ("Segoe UI", 12, "bold")
    FONTE_OPCAO = ("Segoe UI", 11)
    FONTE_BOTAO = ("Segoe UI", 11, "bold")
    FONTE_RODAPE = ("Segoe UI", 9)

    # ============================================================
    # FUNÇÕES
    # ============================================================

    def criar_botoes_de_opcao(root, botoes, op):

        frame = tk.Frame(root)
        frame.pack(fill="x", padx=30, pady=(5, 15))

        for i, texto in enumerate(botoes):

            rb = tk.Radiobutton(
                frame,
                text=texto,
                variable=op,
                value=i + 1,
                font=FONTE_OPCAO,
                anchor="w",
                padx=10,
                pady=7,
                cursor="hand2"
            )

            rb.pack(fill="x")


    def gerar_janela(elementos):

        ROOT = tk.Tk()

        ROOT.title("Gerador de Termos")
        ROOT.geometry(f"{LARGURA}x{ALTURA}")
        ROOT.resizable(False, False)

        # --------------------------------------------------------
        # HEADER
        # --------------------------------------------------------

        header = tk.Frame(ROOT)
        header.pack(fill="x", padx=35, pady=(25, 10))

        titulo = tk.Label(
            header,
            text="GERADOR DE TERMOS",
            font=FONTE_TITULO
        )
        titulo.pack(side="left")

        protocolo = tk.Label(
            header,
            text="N° PROTOCOLO: ____",
            font=FONTE_PROTOCOLO
        )
        protocolo.pack(side="right")

        # Linha separadora
        ttk.Separator(
            ROOT,
            orient="horizontal"
        ).pack(fill="x", padx=35, pady=(5, 25))

        # --------------------------------------------------------
        # ÁREA PRINCIPAL
        # --------------------------------------------------------

        conteudo = tk.Frame(ROOT)
        conteudo.pack(fill="both", expand=True, padx=35)

        OPCOES = []

        for i in range(len(elementos)):

            titulo_opcao = tk.Label(
                conteudo,
                text=elementos[i][0],
                font=FONTE_LABEL,
                justify="left",
                anchor="w"
            )

            titulo_opcao.pack(
                fill="x",
                pady=(5, 0)
            )

            op = tk.IntVar(value=0)

            criar_botoes_de_opcao(
                conteudo,
                elementos[i][1],
                op
            )

            OPCOES.append(op)

            # Separador entre grupos
            if i < len(elementos) - 1:
                ttk.Separator(
                    conteudo,
                    orient="horizontal"
                ).pack(fill="x", pady=5)

        # --------------------------------------------------------
        # BOTÃO
        # --------------------------------------------------------

        area_botao = tk.Frame(ROOT)
        area_botao.pack(fill="x", padx=35, pady=(10, 15))

        enviar = tk.Button(
            area_botao,
            text="Enviar",
            width=20,
            height=2,
            font=FONTE_BOTAO,
            cursor="hand2",
            command=ROOT.destroy
        )

        enviar.pack()

        # --------------------------------------------------------
        # FOOTER
        # --------------------------------------------------------

        footer = tk.Label(
            ROOT,
            text="made by mrl.",
            font=FONTE_RODAPE
        )

        footer.pack(pady=(0, 12))

        ROOT.mainloop()

        return OPCOES

    # ============================================================
    # MENU INICIAL
    # ============================================================

    opcao_inicial = gerar_janela([
        [
            "Bem-vindo! Selecione uma opção:",
            (
                "1. Gerar termos aditivos",
                "2. Gerar Relatório",
                "3. Gerar Portaria",
                "4. Atualizar Número de Protocolos",
                "5. Encerrar"
            )
        ]
    ])

    dados_retorno["opcao"] = opcao_inicial[0].get()

    # ============================================================
    # TERMOS
    # ============================================================

    if dados_retorno["opcao"] == 1:

        opcoes = gerar_janela([
            [
                "Bem-vindo ao Gerador de Termos!\n"
                "Escolha o tipo de termo que deseja gerar:",
                ("Contrato", "Ata")
            ],

            [
                "Deseja analisar a planilha e gerar os documentos imediatamente?",
                ("Sim", "Não")
            ]
        ])

        dados_retorno["dados_extras"] = {
            "tipo_termo": opcoes[0].get(),
            "atualizar_dados": "n" if opcoes[1].get() == 2 else "s",
            "analise_imediata": "n" if opcoes[2].get() == 2 else "s"
        }

    # ============================================================
    # RELATÓRIO
    # ============================================================

    if dados_retorno["opcao"] == 2:

        opcoes = gerar_janela([
            [
                "Bem-vindo ao Gerador de Relatórios!\n"
                "Deseja gerar relatório para qual tipo de termo?",
                ("Contrato", "Ata", "Ambos")
            ]
        ])

        dados_retorno["dados_extras"] = {
            "tipo_relatorio": opcoes[0].get()
        }

    return dados_retorno
    
MAIN()
