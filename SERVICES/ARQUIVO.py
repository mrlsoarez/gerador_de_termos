from openpyxl import load_workbook

from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

from ENV.environment import pegar_modelos
from SERVICES.Classes.Documento import Documento, Termo, Relatorio
from SERVICES.Classes.Verificador import Verificador 

import os
import shutil
from docx import Document
from docx2pdf import convert

from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Pt

RELATORIO_INFO = []

modelo_relatorio = pegar_modelos("relatorio")

# she was so happy to look a mess
# affs
def PROCESSAR_ARQUIVO(sheet_planilha, gerenciador, op):
    
    def gerar_termos(sheet, ordem):
                
        verificador = Verificador("", sheet)   
        gerenciador.entrarEmPasta("WORD")
        
        def verificar_informacoes_iniciais():
            verificacao = verificador.checar_campos_planilha()
            
            if (verificacao["resultado"]):
                verificador.set_info()
                if (not gerenciador.verificarArquivo(verificador.arq)):
                    doc = Documento(verificador.contratado)
                    doc.set_af(verificador.af)
                    doc.set_arq_nome(verificador.arq)
                    return doc
                return False
            else: 
                print(verificacao["mensagem"])
                return False 
                
        def colher_informacoes_termo(doc):
            
            def customizar_mensagem(tipo):
                if (tipo == "Locação"): return "Por este instrumento, em caráter DEFINITIVO, atestamos que a locação acima identificada atende às exigências contratuais."
                return f"Por este instrumento, em caráter DEFINITIVO, atestamos que os {tipo.lower()} acima identificados atendem às exigências contratuais."
            
            def definir_gestor(tipo):
                if (tipo == "contrato"):
                    return "RONALDO DE SOUZA MARCÍLIO\nGESTOR DE CONTRATOS"   
                else:
                    return "MURILO SOARES DE OLIVEIRA\nGESTOR DE ATAS"
            
            mapa = verificador.mapa 
            
            mensagem = customizar_mensagem(str(sheet[mapa['tipo_nota']].value))
            endereco = rf"{gerenciador.pasta_atual}\{doc.arq}"
            gestor = definir_gestor(gerenciador.tipo_arquivo)
            contrato = sheet[mapa['n_contrato']].value
            modelo_termo = pegar_modelos("termo")
            objeto = sheet[mapa['objeto']].value 
            tipo = gerenciador.tipo_arquivo
            contratado = doc.contratado
            numero_af = doc.af 
            
            return Termo(contratado, endereco, ordem, contrato, objeto, numero_af, mensagem, gestor, tipo, modelo_termo)
    
        doc_verificacao = verificar_informacoes_iniciais()
        
        if (doc_verificacao != False):
            termo = colher_informacoes_termo(doc_verificacao)
            termo.criar_arquivo()    
            termo.salvar_pdf()    
    
    def gerar_protocolo(sheet):
        
        verificador = Verificador("", sheet)   
        RELATORIO = []
        
        def verificar_informacoes_iniciais():
            verificacao = verificador.checar_campos_planilha()
            doc = ""
            if (verificacao["resultado"]):
                doc = Documento(verificador.contratado)
            else:
                print(verificacao["mensagem"])
            return verificacao["resultado"], doc 
        
        def colher_informacao_relatorio(doc):
                    
            mapa = verificador.mapa         
                        
            contratado = doc.contratado
            liquidacao = sheet[mapa['numero_liquidacao']].value
            data = sheet[mapa['data_liquidacao']].value 
            valor = sheet[mapa['valor_bruto_liquidacao']].value 
            modelo_termo = pegar_modelos("protocolo")
            
            return Relatorio(contratado, liquidacao, data, valor, modelo_termo)
        
        def copiar_relatorio():
            protocolo = rf'{gerenciador.pasta_atual}\REMESSA X - PROTOCOLO N° {gerenciador.numero_protocolo}.docx'
            print(protocolo, gerenciador.verificarArquivo(protocolo))
                    
            
        verificacao = verificar_informacoes_iniciais()
        
        if (verificacao[0]):
            doc = verificacao[1]
            relatorio = colher_informacao_relatorio(doc)
            copiar_relatorio()
          
            #RELATORIO_INFO.append(colher_informacao_relatorio(doc))
        pass            
    
    PLANILHA = load_workbook(sheet_planilha, data_only= True)


    for ordem in (PLANILHA.sheetnames):
        if (ordem.isdigit() and op == "1"):
            gerar_termos(PLANILHA[ordem], ordem)
        elif (ordem.isdigit() and op == "2"):
            gerar_protocolo(PLANILHA[ordem])
    
    
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
            
        

        