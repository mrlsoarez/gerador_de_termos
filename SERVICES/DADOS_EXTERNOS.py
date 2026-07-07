from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 

URL = "http://bataguassums.biosnet.com.br:8079/transparencia/"

def COLETAR_DADOS_EXTERNOS(direcao, link, tipo):

    if tipo == "MODO_WEB":
        try: 
            driver = webdriver.Chrome()

            # entra na raiz do sistema
            driver.get("http://bataguassums.biosnet.com.br:8079/transparencia/")

            # aguarda a página carregar
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )

            # agora navega para despesas
            driver.execute_script(direcao)

            driver.get(link)
            
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            # procura o botão
            botao = WebDriverWait(driver, 25).until(
                EC.element_to_be_clickable((By.ID, "btnExportarXLS"))
            )

            botao.click()
            time.sleep(10)
        except:
            print(f"Não foi possível capturar a planilha referente a... {link}. Permanecendo com os dados anteriores.")
    elif tipo == "MODO_JSON": 
        pass 

def COLETAR_DADOS_EXTERNOS_JSON():
    def get_contratos(): 
        return { "parametros": {
                    "ConectarExercicio": "2026",
                    "Listagem": "Contratos",
                    "Ano": "2026",
                    "Empresa": "1",
                    "MostraDadosConsolidado": "True",
                    "ContratosApenasPublicados": "False"
                    },   
                "dados_para_planilha": {
                    "Processo": "PROCLIC", 
                    "Ano": "ANO",
                    "Contratado": "FORNECEDOR",
                    "Modalidade": "LICIT",
                    "Número Modalidade": "NUMLICMOD",
                    "Contrato": "CONTRATONUM",
                    "Objeto": "OBJETO_COMPLETO",
                    "Vigência": "VIGENF",
                    "Fiscais": "RESPON"
                },
                "tipo": "contratos"
            }

    C = "/transparencia/VersaoJson/LicitacoesEContratos/"
import requests
from openpyxl import load_workbook


session = requests.Session()

session.get("")


def buscar_dados(url, json):
    response = session.get(url, params=json["parametros"])
    dados = response.json()
    dados_tratados = realizar_tratativa_nos_dados(json["tipo"], json["dados_para_planilha"], dados)
    return dados_tratados

#dto
def realizar_tratativa_nos_dados(tipo, estrutura_planilha, dados):            
    
    def formatar_numero_contrato(dict):
        if (int(dict['Ano']) >= 2025):
            parse = f"{dict['Contrato'][2:]}/{dict['Ano'][2:]}"
            dict['Contrato'] = parse
            
    dados_tratados = []
    for index in range(len(dados)):
        dict = {}
        for chave in estrutura_planilha:       
            DADO = dados[index][estrutura_planilha[chave]]
            dict[chave] = DADO
        
        if (tipo == "contratos"):
            formatar_numero_contrato(dict) 
        dados_tratados.append(dict)
    
    return dados_tratados



"""========================================"""

def salvar_planilha(dados, caminho):
    # temp 
    
    def limpar_planilha(sheet):
        for row in sheet.iter_rows():
            for cell in row:
                cell.value = None
                
    PLANILHA = load_workbook(caminho)
    SHEET = PLANILHA.active
    
    limpar_planilha(SHEET)
  
    HEAD = []
    for keys in dados[0]:
        HEAD.append(keys)
    SHEET.append(HEAD)
    
    for linha in dados: 
        SHEET.append([linha.get(coluna, "") for coluna in HEAD])
    print(dados)
    PLANILHA.save(caminho)


#salvar_planilha(dados_contrato, r"C:\Users\Usuario\Documents\MRL\1. ANÁLISE DE PAGAMENTOS\INFO\BASE\ANÁLISE FISCAL.xlsx")
        
"""
def get_json_despesas():
    return  { "parametros": {
            "ConectarExercicio": "2026",
            "Listagem": "DespesasGerais",
            "DiaInicioPeriodo": "01",
            "MesInicialPeriodo": "01",
            "DiaFinalPeriodo": "31",
            "MesFinalPeriodo": "12",
            "Ano": "2026",
            "Empresa": "1",
            "MostrarFornecedor": "True",
            "MostraDadosConsolidado": "False",
            "UFParaFiltroCOVID": "",
            "MostrarCNPJFornecedor": "True",
            "ApenasIDEmpenho": "False",
            }, 
            "celula_excel": {
                "NOMEFOR": "",
                "CODIGO": "",
                "TPEM": "",
                "DATAE": "",
                "NOMEFOR": "",
                "PROJETO_ATIVIDADE_NOME": "",
                "PROGRAMANOME": "",
                "FICHA": "",
                "FONTE_STN": "",
            }
        }
    
def get_liq_despesas():
     return  {
        "parametros": {
           "ConectarExercicio": "2026",
            "Listagem": "EmpenhosDespesas_Liquidado_PorNumeroEmpenho",
            "intNumeroEmpenho": "NumeroEmpenho",
            "strTipoEmpenho": "TipoEmpenho",
            "DiaInicioPeriodo": "01",
            "MesInicialPeriodo": "01",
            "DiaFinalPeriodo": "31",
            "MesFinalPeriodo": "12",
            "Ano": "2026",
            "Empresa": "1",
            "IDButton": "lnkDespesasPor_NotaEmpenho",
            "MostrarFornecedor": "True",
            "MostraDadosConsolidado": "False",
        },
        "celula_excel": {
            "a": ""
        } 
    }

def get_json_contratos():
    pass 

def pegar_informacoes(url, json):
    response = session.get(url, params=json["parametros"])
    dados = response.json()
    
    dados_para_planilha = []
    dict = {}
    
    parametro_dados_planilha = deepcopy(json["celula_excel"])

    for index in range(len(dados)):
        dict = {}
        for chave in parametro_dados_planilha:
            dict[chave] = dados[index][chave]
        dados_para_planilha.append(dict)
 
    
    print(dados_para_planilha[1]) 
       
    for index2 in range(len(dados_para_planilha)):
        break
        #print(dados_para_planilha[index2])
    


empenhos = pegar_informacoes("http://bataguassums.biosnet.com.br:8079/transparencia/VersaoJson/Despesas/", get_liq_despesas())
"""
#pegar_informacoes()

