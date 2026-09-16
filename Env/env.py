
import os 

from openpyxl import load_workbook
from pathlib import Path

from Services.Calendario import EncontrarData 

# Variável de preparação do ambiente inicial, criando pastas

class Env: 
        
    def __init__(self):
        self.pasta_base = rf"{Path.home()}\Documents\MRL2"
        self.pasta_planilha = rf"{self.pasta_base}\MODELOS BASE\1. PLANILHAS - ANÁLISE FISCAL - ATA & CONTRATOS"
        self.pasta_protocolo = rf"{self.pasta_base}\MODELOS BASE\2. CONTROLE PROTOCOLO"
        self.pasta_atual = self.pasta_base 
        pass 
    
    def criarPasta(self, pasta, navegar = False):
        os.makedirs(pasta, exist_ok = True)
        if (navegar): self.pasta_atual = pasta
        
    def criarPastasIniciais(self):
        self.criarPasta(self.pasta_base)
        self.criarPasta(self.pasta_planilha)
        self.criarPasta(self.pasta_protocolo)
    
    def criarPastasTermos(self): 
        self.criarPasta(rf"{self.pasta_atual}\1. ANÁLISE DE PAGAMENTOS", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\{EncontrarData("ano")}", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\{EncontrarData('mes', False)}. {EncontrarData('mes', True)}", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\{EncontrarData('dia')}-{EncontrarData('mes', False)}", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\REMESSA - PROTOCOLO {self.numero_protocolo}", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\PROTOCOLOS")
        self.criarPasta(rf"{self.pasta_atual}\TERMOS", navegar = True)
        self.criarPasta(rf"{self.pasta_atual}\ATAS")
        self.criarPasta(rf"{self.pasta_atual}\CONTRATOS")

    def setPlanilha(self):
        try: 
            p = load_workbook(rf"{self.pasta_planilha}\ANÁLISE FISCAL.xlsx")
        except: 
            print("-> Planilha não encontrada!")
        else: 
            self.planilha = p
            print(f"-> Planilha de análise fiscal encontrada: {self.planilha}")
    
    def setProtocolo(self):
        try: 
            TXT_PROTOCOLO = rf"{self.pasta_protocolo}\protocolo.txt"
        except:
            print("-> Protocolo não encontrado!")
        else: 
            print("-> Protocolo encontrado!")
            string = ""
            with open(TXT_PROTOCOLO, "r") as txt:
                string = txt.read()
                txt.close()
                self.numero_protocolo = string 

def inicializarAmbiente():
    controladorAmbiente = Env()
    controladorAmbiente.criarPastasIniciais()
    controladorAmbiente.setProtocolo()
    controladorAmbiente.criarPastasTermos()
    controladorAmbiente.setPlanilha()
    return controladorAmbiente

