import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time
from openpyxl import load_workbook

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

def ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(localizacao_planilha):
    
    pasta = r"C:\Users\mrl\Downloads"

    def converter_para_xlsx(arquivo):
        arquivo_xls = arquivo
        arquivo_xlsx = arquivo_xls + "x"
        df = pd.read_excel(arquivo_xls, engine="xlrd")
        df.to_excel(arquivo_xlsx, index=False)
        os.remove(arquivo_xls)

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
    
    try:
        converter_para_xlsx(rf"{pasta}\Portal Transparencia Despesas Gerais - Exercício 2026.xls")
        converter_para_xlsx(rf"{pasta}\Portal Transp. Despesas Liquidadas.xls")
    except: 
        pass

    ANALISE_FISCAL = load_workbook(localizacao_planilha)
    planilha_empenhos = load_workbook(rf"{pasta}\Portal Transparencia Despesas Gerais - Exercício 2026.xlsx")
    planilha_liquidacao = load_workbook(rf"{pasta}\Portal Transp. Despesas Liquidadas.xlsx")

    try:
        atualizar_planilha(planilha_empenhos, "Empenhos")
        atualizar_planilha(planilha_liquidacao, "Liquidacoes")
    except:
        print("Não foi possível completar a substituição dos dados na planilha.")
    else:
        ANALISE_FISCAL.save(localizacao_planilha)

def INICIAR_PLANILHA(localizacao_planilha):

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(localizacao_planilha)

    print("Ouvindo planilha... aperte qualquer botão para encerrar e fechar a planilha.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Encerrando...")
    finally:
        wb.Close(SaveChanges=True)  # fecha a planilha
        excel.Quit()                 # fecha o Excel
