from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, pegar_numero_protocolo
from SERVICES.PLANILHA import ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS, INICIAR_PLANILHA
from SERVICES.DADOS_EXTERNOS import COLETAR_DADOS_EXTERNOS
from MODULES.GerenciarArquivos import GerenciarArquivos


# lista de dependencias sao elas
#win32
#selenium
#openpyxl
#pandas -- xlrd
r"""

import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from openpyxl import load_workbook
import time
from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, pegar_numero_protocolo

PLANILHA = pegar_planilha_termo("ANÁLISE FISCAL - ATA.xlsx")

def COLETAR_DADOS_EXTERNOS(direcao, link):

    try: 
        driver = webdriver.Chrome()

        # entra na raiz do sistema
        driver.get("http://bataguassums.biosnet.com.br:8079/transparencia/")

        # aguarda a página carregar
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        # agora navega para despesas
        driver.execute_script(direcao)

        driver.get(link)
        
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        # procura o botão
        botao = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.ID, "btnExportarXLS"))
        )

        botao.click()
        time.sleep(15)
    except:
        print(f"Não foi possível capturar a planilha referente a... {link}. Permanecendo com os dados anteriores.")

def ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS():
    
    pasta = r"C:\Users\mrl\Downloads"

    def converter_para_xlsx(arquivo):
        try:
            arquivo_xls = arquivo
            arquivo_xlsx = arquivo_xls + "x"
            df = pd.read_excel(arquivo_xls, engine="xlrd")
            df.to_excel(arquivo_xlsx, index=False)
            os.remove(arquivo_xls)
        except: 
            pass

    def atualizar_planilha(sheet_origem, sheet_destino):
        origem = sheet_origem.active
        destino = ANALISE_FISCAL[sheet_destino]
       
        for linha_idx, linha in enumerate(origem.iter_rows(values_only=True), start=1):
            for coluna_idx, valor in enumerate(linha, start=1):
                destino.cell(
                    row=linha_idx,
                    column=coluna_idx,
                    value=valor
                )
       
    converter_para_xlsx(rf"{pasta}\Portal Transparencia Despesas Gerais - Exercício 2026.xls")
    converter_para_xlsx(rf"{pasta}\Portal Transp. Despesas Liquidadas.xls")

    ANALISE_FISCAL = load_workbook(PLANILHA)

    planilha_empenhos = load_workbook(rf"{pasta}\Portal Transparencia Despesas Gerais - Exercício 2026.xlsx")
    planilha_liquidacao = load_workbook(rf"{pasta}\Portal Transp. Despesas Liquidadas.xlsx")

    try:
        atualizar_planilha(planilha_empenhos, "Empenhos")
        atualizar_planilha(planilha_liquidacao, "Liquidacoes")
    except:
        print("Não foi possível completar a substituição dos dados na planilha.")
    else:
        ANALISE_FISCAL.save(r"C:\Users\mrl\Documents\teste\ANÁLISE FISCAL - ATA.xlsx")

def INICIAR_PLANILHA():

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(pegar_planilha_termo("ANÁLISE FISCAL - ATA.xlsx"))

    print("Ouvindo planilha... aperte qualquer botão para encerrar e fechar a planilha.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Encerrando...")
    finally:
        wb.Close(SaveChanges=True)  # fecha a planilha
        excel.Quit()                 # fecha o Excel

# cerca de 1 min pra rodar

ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS()

"""

"""
///***********//////////
"""


def MAIN():
    
    #pergunta pra saber se eh ata/contrato
    """
    pergunta_inicial = CAPTURAR_RESPOSTA(
        "Bem vindo ao gerador de termos! Escolha dentre as opções para gerar: \n1. Contrato\n2. Ata\n-> ", ("1", "2")
    )

    #pergunta pra saber se quer pegar os dados novos de empenho e liq
    pergunta_update = CAPTURAR_RESPOSTA(
        "Deseja atualizar os dados da planilha? (S/N) -> ", ("s", "n")
    )
    """
    
    
    # A coleta de dados externos oferece uma estrutura que envolve
        # -> Atualização da planilha contrato com dados atualizados dos contratos
            # -> Inclusão dos empenhos e liquidações atualizadas
            # -> Inclusão dos servidores
        ## Futuramente, inclusão de JSON com os dados de ata
        
    pergunta_inicial = "1"
    pergunta_update = "s"
    r"""
    
    
    GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), TIPO_TERMO["tipo"])
    GERENCIADOR_PASTAS.criar_pasta_termos()
    
    """
    TIPO_TERMO = pegar_tipo_termo(pergunta_inicial)
    PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO["arquivo"])
    
    if (pergunta_update == "s"): 
        
        DADOS = []       
         
        dados_contratos = COLETAR_DADOS_EXTERNOS("contratos")
        dados_servidores =  COLETAR_DADOS_EXTERNOS("servidores")
        dados_empenhos = COLETAR_DADOS_EXTERNOS("empenhos")
        dados_liquidacao =  COLETAR_DADOS_EXTERNOS("liquidacoes")
        
        #DADOS.append(dados_contratos)
        #DADOS.append(dados_servidores)
        
        ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Base", dados_contratos)
        ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Servidores", dados_servidores)
        ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Empenhos", dados_empenhos)
        ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Liquidacoes", dados_liquidacao)

        return

    #INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS)

def CAPTURAR_RESPOSTA(mensagem, dado_esperado):

    PERGUNTA = input(mensagem).lower()

    while PERGUNTA not in dado_esperado:
        print(("Opção não válida, por favor, digite uma das opções ao lado ", dado_esperado))
        PERGUNTA = input(mensagem).lower()
    return PERGUNTA

MAIN()