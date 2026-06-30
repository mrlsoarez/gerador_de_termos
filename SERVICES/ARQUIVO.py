from openpyxl import load_workbook
from MODULES.GerenciarArquivos import GerenciarArquivos
from ENV.environment import pegar_modelos
import os
import os
import shutil

MAPEAMENTO = {
    "contratado": "B4",
    "n_contrato": "E4",
    "objeto": "F4",
    "numero_empenho": "A8",
    "numero_liquidacao": "A12",
    "data_liquidacao": "B12",
    "valor_bruto_liquidacao": "C12",
    "tipo_nota": "D16",
    "numero_af": "A20",
}

RELATORIO_INFO = []

class Documento: 
    
    endereco_protocolo = pegar_modelos("protocolo")
    
    def __init__(self, contratado):
        self.contratado = contratado
    
    def copiar_arquivo(self, antigo, novo):
        shutil.copy(antigo, novo)
        
class Termo(Documento):
    
    def __init__(self, contratado, ordem, contrato, objeto, af, mensagem, gestor):
        super().__init__(contratado)
        self.ordem = ordem
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.gestor = gestor 

        
    def criar_arquivo(self):
        pass 
     
class Relatorio(Documento):
    
    def __init__(self, contratado, liquidacao, valor, data):
        super().__init__(contratado)
        self.liquidacao = liquidacao
        self.valor = valor
        self.data = data 
        
def PROCESSAR_ARQUIVO(sheet_planilha, gerenciador):
    
    PLANILHA = load_workbook(sheet_planilha, data_only= True)
    
    gerenciador.entrar_na_pasta("WORD")
    gerenciador_modulo_arquivos = GerenciarArquivos(gerenciador.pasta_atual, None)
    print(gerenciador_modulo_arquivos.pasta_base)
    
    def instanciar_documento(sheet):
        return Documento(sheet[MAPEAMENTO['contratado']].value)
    
    def verificar_se_arquivo_existe(contratado, sheet):       
        numero_af = sheet[MAPEAMENTO['numero_af']].value
        arq_prototipo = f'{nome}. {contratado} - AF {numero_af[:4]}.docx'      
        if (gerenciador.verificar_arquivo(arq_prototipo)): return True   
        return arq_prototipo
        
    def colher_informacoes_relatorio(contratado, sheet):
        liq = sheet[MAPEAMENTO['numero_liquidacao']].value 
        val = sheet[MAPEAMENTO['valor_bruto_liquidacao']].value
        data = sheet[MAPEAMENTO['data_liquidacao']].value 
            
        rel = Relatorio(contratado, liq, val, data)
        return rel
        
    def colher_informacoes_termo(ordem, contratado, sheet): 
        def customizar_mensagem(tipo):
            if (tipo == "Locação"): return "Por este instrumento, em caráter DEFINITIVO, atestamos que a locação acima identificada atende às exigências contratuais."
            return f"Por este instrumento, em caráter DEFINITIVO, atestamos que os {tipo.lower()} acima identificados atendem às exigências contratuais."
        numero_af = sheet[MAPEAMENTO['numero_af']].value
        contrato = sheet[MAPEAMENTO['n_contrato']].value 
        objeto = sheet[MAPEAMENTO['objeto']].value 
        mensagem = customizar_mensagem(str(sheet[MAPEAMENTO['tipo_nota']].value))
        return Termo(contratado, ordem, contrato, objeto, numero_af, mensagem, "wanda")
    
    for ordem in PLANILHA.sheetnames:
        try: 
            nome = int (ordem)
        except: 
            pass 
        else: 
            pass
        
            doc = instanciar_documento(PLANILHA[ordem])
            RELATORIO_INFO.append(colher_informacoes_relatorio(doc.contratado, PLANILHA[ordem]))
            arquivo_termo = verificar_se_arquivo_existe(doc.contratado, PLANILHA[ordem])
            
            if (arquivo_termo != True):
                termo = colher_informacoes_termo(ordem, doc.contratado, PLANILHA[ordem])
                """
                ****** end_modelo needs to be in class
                """
                end_modelo = pegar_modelos("termo")
                end_novo = rf"{gerenciador.pasta_atual}\{arquivo_termo}"
                #termo.copiar_arquivo(end_modelo, end_novo)
                """
                ******
                """
                pass
            
            
    
        