
import os 
from MODULES.EncontrarData import EncontrarData
from ENV.environment import pegar_numero_protocolo

class GerenciarArquivos:
    
    numero_protocolo = pegar_numero_protocolo()
    pasta_atual = None 
    
    def __init__(self, pasta_base, tipo_arquivo):
        self.pasta_base = pasta_base
        self.tipo_arquivo = tipo_arquivo
        
    def criar_pasta(self, nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    def entrar_na_pasta(self, nome_pasta):
        os.chdir(nome_pasta)
        self.pasta_atual = rf"{self.pasta_atual}\{nome_pasta}"
        
    def verificar_arquivo(self, arq):
        return os.path.exists(rf"{self.pasta_atual}/{arq}")
    
    def criar_pasta_termos(self):
        
        mes_numero = EncontrarData("mes", False)
        mes_extenso = EncontrarData("mes", True) 
        dia = EncontrarData("dia")
        
        os.chdir(self.pasta_base)

        PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
        PASTA_DIA = f"{dia}-{mes_numero}"
        PROTOCOLO = f"REMESSA X - PROTOCOLO N° {self.numero_protocolo}"
        
        self.criar_pasta(PASTA_MES, True)
        self.criar_pasta(PASTA_DIA, True)
        self.criar_pasta(PROTOCOLO, True)
        self.criar_pasta(self.tipo_arquivo, True)
        self.criar_pasta("WORD")
        self.criar_pasta("PDF")
        
        self.pasta_atual = rf"{self.pasta_base}\{PASTA_MES}\{PASTA_DIA}\{PROTOCOLO}\{self.tipo_arquivo}"
        
        
    
        