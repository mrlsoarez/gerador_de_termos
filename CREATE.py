
from MODULES.object_creator.create_object import capturar_info_planilha, Termo
from MODULES.date_dealer.date import pegar_data_hoje_ptbr
from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, pegar_numero_protocolo

import shutil 
import os 

NUMERO_ATUAL_PROTOCOLO = pegar_numero_protocolo()

def MAIN():

    PASTA_BASE = pegar_endereco_base()
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
    CRIAR_PASTAS_INICIAIS(PASTA_BASE, TIPO_TERMO)

    GERAR_TERMOS_E_PROTOCOLOS(TIPO_TERMO)

# Cria as pastas para guardar os termos e relatório
def CRIAR_PASTAS_INICIAIS(endereco_base, tipo_termo):

    def criar_pasta(nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    print(f"Criando as pastas iniciais caso não existam... o protocolo dessa remessa é {NUMERO_ATUAL_PROTOCOLO}")
    
    PASTA_DATA_DE_HOJE = rf"{endereco_base}\{pegar_data_hoje_ptbr("-")[:5]}"
    PASTA_PROTOCOLO = f"REMESSA X - PROTOCOLO {NUMERO_ATUAL_PROTOCOLO}"

    PLANILHA_PASTA_ORIGINAL = pegar_planilha_termo(tipo_termo['arquivo'])
    PLANILHA_NOVA_PASTA = rf"{PASTA_DATA_DE_HOJE}\{PASTA_PROTOCOLO}\{tipo_termo['arquivo']}"

    criar_pasta(PASTA_DATA_DE_HOJE, navegar = True)
    criar_pasta(PASTA_PROTOCOLO, navegar = True)
  
    shutil.copyfile(PLANILHA_PASTA_ORIGINAL, PLANILHA_NOVA_PASTA)

    criar_pasta("PROTOCOLOS")
    criar_pasta("TERMOS", navegar = True) 
    criar_pasta("PDF")
    criar_pasta("WORD", navegar = True)

def GERAR_TERMOS_E_PROTOCOLOS(tipo, mesmo_protocolo = False):

    localizacao_planilha = pegar_planilha_termo(tipo["arquivo"])
    
    TERMOS = capturar_info_planilha(localizacao_planilha)
    
    for index in range(len(TERMOS)):
        TERMOS[index].criar_termo(tipo)
        #print(f"Termo criado!!! (verificar na pasta) --> {TERMOS[index].contratado} (AF {TERMOS[index].af})")
        pass 
    
    os.chdir(r"../../PROTOCOLOS")

    Termo.criar_relatorio(TERMOS, NUMERO_ATUAL_PROTOCOLO, mesmo_protocolo)
   
    #print("Relatório gerado! :)")


MAIN()