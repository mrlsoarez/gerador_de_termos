from openpyxl import load_workbook

from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

from ENV.environment import pegar_modelos
from SERVICES.Classes.Documento import Documento

import os
import shutil
from docx import Document
from docx2pdf import convert

from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Pt




RELATORIO_INFO = []

modelo_termo = pegar_modelos("termo")
modelo_relatorio = pegar_modelos("relatorio")

# she was so happy to look a mess
# affs
def PROCESSAR_ARQUIVO(sheet_planilha, gerenciador, op):
    
    # trabalha o relatorio e o termo aqui
    def gerar_termos(sheet):
            
            # todos os dados da planilha precisam estar preenchidos para checar se arquivo existe
            def verificar_se_arquivo_existe():
                
                doc_intermediario = Documento("")
                mapa = doc_intermediario.get_mapeamento_termos()
                
                verificao = doc_intermediario.checar_campos_planilha(sheet)
                
                if (verificao[0] == False):
                    print(f'{verificao[1]}')
                    return 
                
                contratado = sheet[mapa['contratado']].value
                numero_af = sheet[mapa['numero_af']].value[:4]
                
                doc_intermediario.set_contratado(contratado)
                return gerenciador.verificarArquivo(f"{contratado} - AF {numero_af}.docx"), doc_intermediario
                               
            # 1. PRIMEIRO ABRE A PLANILHA E AVALIA SE A SHEET É UM NUMERO, SE FOR INICIA A ANÁLISE ok
            # 2. COMEÇA VERIFICANDO SE O ARQUIVO EXISTE NA PASTA MONTANDO UMA VARIÁVEL INTERMEDIÁRIA E TEMPORÁRIA ok
            # 3. SE NÃO EXISTIR, UTILIZAR O MAPEAMENTO PARA VERIFICAR SE AS INFORMAÇÕES VITAIS ESTÃO PREENCHIDAS
            # 4. SE O MAPEAMENTO ESTIVER OK, INSTANCIAR OBJETO E INICIALIZAR O DOCUMENTO
            
            gerenciador.entrarEmPasta("WORD")
            
            verificacao_inicial = verificar_se_arquivo_existe()
            
            if (verificacao_inicial[0] == False):
                DOC = verificacao_inicial[1]
                print(DOC)
                           
    PLANILHA = load_workbook(sheet_planilha, data_only= True)

    if (op == "1" or op == "2"):
        for ordem in (PLANILHA.sheetnames):
            if (ordem.isdigit() and op == "1"):
                gerar_termos(PLANILHA[ordem])
         
    
    
    """
    
    
    
    
    def instanciar_documento(sheet):
        return Documento(sheet[MAPEAMENTO['contratado']].value)
    
    def verificar_se_arquivo_existe(contratado, sheet):       
        numero_af = sheet[MAPEAMENTO['numero_af']].value
        arq_prototipo = f'{nome}. {contratado} - AF {numero_af[:4]}.docx'      
        if (gerenciador.verificarArquivo(arq_prototipo)): return True   
        return arq_prototipo
        
    def colher_informacoes_relatorio(contratado, sheet):
        liq = sheet[MAPEAMENTO['numero_liquidacao']].value 
        val = sheet[MAPEAMENTO['valor_bruto_liquidacao']].value
        data = sheet[MAPEAMENTO['data_liquidacao']].value 
            
        rel = Relatorio(contratado, liq, val, data)
        return rel
        
    def colher_informacoes_termo(ordem, contratado, sheet, arq): 
        def customizar_mensagem(tipo):
            if (tipo == "Locação"): return "Por este instrumento, em caráter DEFINITIVO, atestamos que a locação acima identificada atende às exigências contratuais."
            return f"Por este instrumento, em caráter DEFINITIVO, atestamos que os {tipo.lower()} acima identificados atendem às exigências contratuais."
        endereco = rf"{gerenciador.pasta_atual}\{arq}.docx"
        numero_af = sheet[MAPEAMENTO['numero_af']].value
        contrato = sheet[MAPEAMENTO['n_contrato']].value 
        objeto = sheet[MAPEAMENTO['objeto']].value 
        mensagem = customizar_mensagem(str(sheet[MAPEAMENTO['tipo_nota']].value))
        return Termo(contratado, endereco, ordem, contrato, objeto, numero_af, mensagem, gerenciador.tipo_arquivo)
    
    for ordem in PLANILHA.sheetnames:
        try: 
            nome = int (ordem)
        except: 
            pass 
        else: 
            pass
        
            
            RELATORIO_INFO.append(colher_informacoes_relatorio(doc.contratado, PLANILHA[ordem]))
            nome_arquivo = verificar_se_arquivo_existe(doc.contratado, PLANILHA[ordem])
            #nome_arquivo = "1. CENTRO AMERICA FROTAS LTDA - AF 1950"
            # Verifica se o arquivo existe, se não cria um termo na pasta atual (dia atual/ remessa x / word)
            if (nome_arquivo != True):
                
                termo = colher_informacoes_termo(ordem, doc.contratado, PLANILHA[ordem], nome_arquivo)
                
                doc = termo.criar_arquivo(modelo_termo)
                doc.save(termo.endereco)
                try:
                    termo.salvar_pdf()
                except: 
                    pass
               
                
                
                ****** end_modelo needs to be in class
                
            
                #termo.copiar_arquivo(end_modelo, end_novo)
                
                ******
            
                pass
    """
            
        

        