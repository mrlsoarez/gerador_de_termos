from openpyxl import load_workbook
import os

MAPEAMENTO = {
    "contratado": "B4",
    "n_contrato": "E2",
    "objeto": "F2",
    "numero_empenho": "A6",
    "numero_liquidacao": "A10",
    "data_liquidacao": "B10",
    "valor_bruto_liquidacao": "C10",
    "tipo_nota": "D14",
    "numero_af": "A20",
}

class Arquivo:
    
    def __init__(self, ordem, contrato, contratado, objeto, af, mensagem, gestor, liquidacao, valor, data):
        self.ordem = ordem
        self.contratado = contratado 
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.gestor = gestor 
        self.liquidacao = liquidacao 
        self.valor = valor 
        self.data = data
        
def PROCESSAR_ARQUIVO(sheet_planilha, gerenciador):
    
    gerenciador.entrar_na_pasta("WORD")
    
    def verificar_se_arquivo_existe():
        PLANILHA = load_workbook(sheet_planilha)
        for nome in PLANILHA.sheetnames:
            try:
                ordem = int(nome)
            except:
                pass
            else: 
                SHEET = PLANILHA[nome]
                contratado = SHEET[MAPEAMENTO['contratado']].value 
                numero_af = SHEET[MAPEAMENTO['numero_af']].value
                arq_prototipo = f'{nome}. {contratado} - AF {numero_af[:4]}.docx'
                if (gerenciador.verificar_arquivo(arq_prototipo)):
                    continue
                
                #novo_termo = Termo(ordem, contrato, contratado, objeto, af, mensagem, gestor)

                
                
    
    verificar_se_arquivo_existe()
    pass