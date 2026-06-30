
from MODULES.object_creator.create_object import capturar_info_planilha, Termo
from MODULES.date_dealer.date import pegar_data_hoje_ptbr
from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, pegar_numero_protocolo

import shutil 
import os 
"""

class CriadorDePastas:
    def __init__(self, pasta_hoje, pasta_protocolo, planilha_original, planilha_copiada):
        self.pasta_hoje = pasta_hoje
        self.pasta_protocolo = pasta_protocolo
        self.planilha_original = planilha_original
        self.planilha_copiada = planilha_copiada
        pass

    def criar_pasta(nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    def mover_planilhas(original, copiada): 
        shutil.copyfile(original, copiada)

def MAIN(incremental = False):

    PASTA_BASE = pegar_endereco_base()
    print(f"O protocolo dessa análise é {NUMERO_ATUAL_PROTOCOLO}.")
    PROMPT_INICIAL = input("Digite o tipo de documentos: (Contrato, ata ou ambos?): ").lower()

    if PROMPT_INICIAL == "ambos":
        TIPO_TERMO = pegar_tipo_termo("contrato")
        CRIAR_PASTAS_INICIAIS(PASTA_BASE, TIPO_TERMO)
        GERAR_TERMOS_E_PROTOCOLOS(TIPO_TERMO)
        TIPO_TERMO = pegar_tipo_termo("ata")
        CRIAR_PASTAS_INICIAIS(PASTA_BASE, TIPO_TERMO)
        GERAR_TERMOS_E_PROTOCOLOS(TIPO_TERMO, mesmo_protocolo = True)
        return 
    
    TIPO_TERMO = pegar_tipo_termo(PROMPT_INICIAL)
    gerenciador = CRIAR_PASTAS_INICIAIS(PASTA_BASE, TIPO_TERMO)
    GERAR_TERMOS_E_PROTOCOLOS(TIPO_TERMO, incremental, gerenciador)

def CRIAR_PASTAS_INICIAIS(endereco_base, tipo_termo):

    def criar_pasta(nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    print(f"Criando as pastas iniciais caso não existam... o protocolo dessa remessa é {NUMERO_ATUAL_PROTOCOLO}")
    
    PASTA_DATA_DE_HOJE = rf"{endereco_base}\{pegar_data_hoje_ptbr("-")[:5]}"
    PASTA_PROTOCOLO = f"REMESSA X - PROTOCOLO {NUMERO_ATUAL_PROTOCOLO}"
    PLANILHA_PASTA_ORIGINAL = pegar_planilha_termo(tipo_termo['arquivo'])
    PLANILHA_NOVA_PASTA = rf"{PASTA_DATA_DE_HOJE}\{PASTA_PROTOCOLO}\{tipo_termo['arquivo']}"

    gerenciador = CriadorDePastas(PASTA_DATA_DE_HOJE, PASTA_PROTOCOLO, PLANILHA_PASTA_ORIGINAL, PLANILHA_NOVA_PASTA)
    
    gerenciador.criar_pasta(PASTA_DATA_DE_HOJE, navegar = True)
    gerenciador.criar_pasta(PASTA_PROTOCOLO, navegar = True)
  
    gerenciador.mover_planilhas(PLANILHA_PASTA_ORIGINAL, PLANILHA_NOVA_PASTA)

    gerenciador.criar_pasta("PROTOCOLOS")
    gerenciador.criar_pasta("TERMOS", navegar = True) 
    gerenciador.criar_pasta("PDF")
    gerenciador.criar_pasta("WORD", navegar = True)
    
    return gerenciador

def GERAR_TERMOS_E_PROTOCOLOS(tipo, incremental = True, mesmo_protocolo = False):

    if (not incremental):
        numero_planilha = input("Deseja começar por uma parte específica? (Y/N): ").lower() 
        if numero_planilha == "y": 
            numero_planilha = input("Insira o número: ")
        else: 
            numero_planilha = False 
        
    localizacao_planilha = pegar_planilha_termo(tipo["arquivo"])

    while (incremental):

        TERMOS = capturar_info_planilha(localizacao_planilha)
        for index in range(len(TERMOS)):
            TERMOS[index].criar_termo(tipo, impressao = True)

        sentinela = input("Deseja ENCERRAR a geração incremental? (Y/N): ").lower()
        if (sentinela == "y"): 
            incremental = False 
            os.chdir(r"../../PROTOCOLOS")
            Termo.criar_relatorio(TERMOS, NUMERO_ATUAL_PROTOCOLO, mesmo_protocolo)

    for index in range(len(TERMOS)):
        TERMOS[index].criar_termo(tipo)
        pass 
    
    os.chdir(r"../../PROTOCOLOS")

    Termo.criar_relatorio(TERMOS, NUMERO_ATUAL_PROTOCOLO, mesmo_protocolo)
   
    #print("Relatório gerado! :)"

def MENU():

    TIPO_DE_GERADOR = input("Digite qual o tipo de geração.:\n1. Incremental\n2. Total\n-> ")

    if (TIPO_DE_GERADOR == "1"): 
        MAIN(incremental = True)
    if (TIPO_DE_GERADOR == "2"):
        MAIN()

MENU()
"""
"""

class GerenciadorDePasta:

    def criar_pasta(self, nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    def settar_nome_planilhas(self, planilha_original, planilha_copiada):
        self.planilha_original = planilha_original 
        self.planilha_copiada = planilha_copiada

    def mover_planilhas(self): 
        shutil.copyfile(self.planilha_original, self.planilha_copiada)
    


gerenciador = GerenciadorDePasta()
NUMERO_ATUAL_PROTOCOLO = pegar_numero_protocolo()

def MAIN(tipo_documento, incremental): 

    def INICIAR(tipo_documento, mesmo_protocolo):
        TIPO_TERMO = pegar_tipo_termo(tipo_documento)
        CRIAR_PASTAS_INICIAIS(TIPO_TERMO)
        GERAR_TERMOS_E_PROTOCOLO(TIPO_TERMO, incremental, mesmo_protocolo)

    def CRIAR_PASTAS_INICIAIS(tipo_termo):

        endereco_base = pegar_endereco_base()

        PASTA_DATA_DE_HOJE = rf"{endereco_base}\{pegar_data_hoje_ptbr("-")[:5]}"
        PASTA_PROTOCOLO = f"REMESSA X - PROTOCOLO {NUMERO_ATUAL_PROTOCOLO}"
        PLANILHA_PASTA_ORIGINAL = pegar_planilha_termo(tipo_termo['arquivo'])
        PLANILHA_NOVA_PASTA = rf"{PASTA_DATA_DE_HOJE}\{PASTA_PROTOCOLO}\{tipo_termo['arquivo']}"

        gerenciador.criar_pasta(PASTA_DATA_DE_HOJE, navegar = True)
        gerenciador.criar_pasta(PASTA_PROTOCOLO, navegar = True)
  
        gerenciador.settar_nome_planilhas(PLANILHA_PASTA_ORIGINAL, PLANILHA_NOVA_PASTA)
        gerenciador.mover_planilhas()

        gerenciador.criar_pasta("PROTOCOLOS")
        gerenciador.criar_pasta("TERMOS", navegar = True) 
        gerenciador.criar_pasta(tipo_termo['corresponde_a'].upper(), navegar = True)
        gerenciador.criar_pasta("PDF")
        gerenciador.criar_pasta("WORD", navegar = True)

    def GERAR_TERMOS_E_PROTOCOLO(tipo_termo, incremental, mesmo_protocolo):

        print(f"Começando a geração de documento... --> {tipo_termo['corresponde_a'].upper()}!! *.✧")
        localizacao_planilha = pegar_planilha_termo(tipo_termo["arquivo"])
        ordem_planilha = False

        if (not incremental):
                perguntar_ordem = input("Deseja começar por uma parte específica? (Y/N): ").lower() 
                if perguntar_ordem == "y": 
                    ordem_planilha = input("Insira o número: ")

        print("Iniciando criação de termo.... *.✧*.✧.*.✧*.✧*.✧*.✧.,*.✧*.✧*.✧*.✧*.✧*.✧*.✧,*.✧*.✧*.✧*.✧*.✧*.✧*.✧.*.✧*.✧*.✧*.✧*.✧*.✧*.✧*.✧*.✧*.✧*.✧")
        def iniciar_geracao(termos, incremental):    

            for index in range(len(termos)):
                print(f"✧ *.✧ -> Criando o seguinte termo... {termos[index].contratado} - AF {termos[index].af} <-- ✧*.✧\n=======================================================================================")
                termos[index].criar_termo(tipo_termo, incremental)

            if (incremental):
                gerenciador.mover_planilhas()
                sentinela = input("Deseja ENCERRAR a geração incremental? (Y/N): ").lower()
                if (sentinela == "y"): 
                    incremental = False

        while True:
            TERMOS = capturar_info_planilha(localizacao_planilha, ordem_planilha)
            iniciar_geracao(TERMOS, incremental)
            if (not incremental): 
                break 
        
        os.chdir(r"../../../PROTOCOLOS")
        Termo.criar_relatorio(TERMOS, NUMERO_ATUAL_PROTOCOLO, mesmo_protocolo)
    
    if (tipo_documento == "ambos"):
        INICIAR("contrato", mesmo_protocolo = False)
        INICIAR("ata", mesmo_protocolo = True)
        return 

    INICIAR(tipo_documento, mesmo_protocolo = False)
    

def MENU():

    print("============================  GERADOR DE TERMO DEFINITIVO =================================")
    TIPO_DE_GERADOR = input("Digite qual o tipo de geração.:\n1. Incremental\n2. Total\n----> ")
    TIPO_DE_DOCUMENTO = input("Digite o tipo de documento que será gerado: (Contrato, ata ou ambos?): ------> ").lower()

    if (TIPO_DE_GERADOR == "1"): 
        incremental = True
    else:
        incremental = False

    MAIN(TIPO_DE_DOCUMENTO, incremental)

MENU()

"""