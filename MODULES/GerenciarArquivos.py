
import os 

class GerenciarArquivos:
    
    def __init__(self, pasta_base):
        self.pasta_base = pasta_base
        
    def criar_pasta(self, nome_pasta, navegar = False):
        os.makedirs(nome_pasta, exist_ok = "True")
        if(navegar): os.chdir(nome_pasta)
    """
    
    def criar_pasta_datas():
        
        mes_numero = EncontrarData("mes", extenso = False)
        mes_extenso = EncontrarData("mes", extenso = True) 
        
        print(mes_numero, mes_extenso)
    """
        