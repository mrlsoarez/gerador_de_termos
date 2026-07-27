
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
    
    def set_pasta_base(self, pasta):
        self.pasta_base = pasta 
        
    def set_tipo_arquivo(self, tipo):
        self.tipo_arquivo = tipo 
        
    def criarPasta(self, nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)

    def entrarEmPasta(self, nome_pasta):
        try:
            self.pasta_atual = rf"{self.pasta_base}\{nome_pasta}"
            os.chdir(self.pasta_atual)
        except: 
            print("Não foi possível entrar na pasta!")
            pass
        
    def verificarArquivo(self, arq):
        return os.path.exists(rf"{self.pasta_atual}/{arq}")
    
   
    
        