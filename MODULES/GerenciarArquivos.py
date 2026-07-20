
import os 

from MODULES.EncontrarData import EncontrarData
from ENV.environment import pegar_numero_protocolo

def verificarArquivo(arq):
    return os.path.exists(rf"{arq}")

class GerenciarArquivos:
    
    numero_protocolo = pegar_numero_protocolo()
    pasta_atual = None 
    
    def __init__(self, pasta_base, tipo_arquivo):
        self.pasta_base = pasta_base
        self.tipo_arquivo = tipo_arquivo
        
    def criarPasta(self, nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    def entrarEmPasta(self, nome_pasta):
        if (os.getcwd() == self.pasta_atual):
            return
        os.chdir(nome_pasta)
        self.pasta_atual = rf"{self.pasta_atual}\{nome_pasta}"
        
    def verificarArquivo(self, arq):
        return os.path.exists(rf"{self.pasta_atual}/{arq}")
    
    def criar_pasta_termos(self):
        
        mes_numero = EncontrarData("mes", False)
        mes_extenso = EncontrarData("mes", True) 
        dia = EncontrarData("dia")
        
        os.chdir(self.pasta_base)

        PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
        PASTA_DIA = f"{dia}-{mes_numero}"
        PROTOCOLO = f"REMESSA X - PROTOCOLO N° {self.numero_protocolo}"
        
        self.criarPasta(PASTA_MES, True)
        self.criarPasta(PASTA_DIA, True)
        self.criarPasta(PROTOCOLO, True)
        self.criarPasta(self.tipo_arquivo, True)
        self.criarPasta("WORD")
        self.criarPasta("PDF")
        
        self.pasta_atual = rf"{self.pasta_base}\{PASTA_MES}\{PASTA_DIA}\{PROTOCOLO}\{self.tipo_arquivo}"
        
        
    
        