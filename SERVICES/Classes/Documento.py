
from openpyxl import load_workbook

from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

from ENV.environment import pegar_modelos

import os
import shutil
from docx import Document
from docx2pdf import convert

from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Pt

class Documento: 

    endereco_protocolo = pegar_modelos("protocolo")
    
    def __init__(self, contratado):
        self.contratado = contratado

    def set_contratado(self, dado):
        self.contratado = dado 
        
    def get_mapeamento_termos(self):
        return {
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
        
    def checar_campos_planilha(self, sheet):
        mapa = self.get_mapeamento_termos()
        for campo, celula in mapa.items():
            valor = sheet[celula].value 
            if (valor == None):
                return False, f'{sheet} -> O campo "{campo}" desta planilha está vazio e sem conteúdo. Necessário preencher e tentar novamente!' 
        return True, ""
    
    def copiar_arquivo(self, antigo, novo):
        shutil.copy(antigo, novo)
        
    def mudar_fonte(texto, name_font):
        texto.font.name = name_font
    
    def mudar_tamanho(texto, px): 
        texto.font.size = Pt(px)

    @staticmethod
    def criar_texto(paragrafo, texto, negrito = False, posicionamento = None, px = None, fonte = None):

        def deixar_negrito(run): 
           run.bold = True 

        def alinhar_texto(paragrafo, alinhado):
            if (alinhado == "Centro"): 
                paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif (alinhado == "Direita"):
                paragrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT 

        run = paragrafo.add_run(str(texto))
        if (negrito): deixar_negrito(run)
        if (posicionamento != None): alinhar_texto(paragrafo, posicionamento)
        if (px != None): Documento.mudar_tamanho(run, px)
        if (fonte != None): Documento.mudar_fonte(run, fonte)

    @staticmethod
    def encontrar_data_de_hoje_em_extenso():  

        dia = EncontrarData("dia")
        mes = EncontrarData("mes", True)
        ano = EncontrarData("ano")
        
        return f"{dia} de {mes} de {ano}"

    @staticmethod
    def adicionar_linha_de_assinatura(paragrafo, assinador): 
        Documento.criar_texto(paragrafo, f"________________________________________\n{assinador}", negrito = True, posicionamento = "Centro", fonte = "Cambria")
    
    @staticmethod
    def modificar_tabela(tabela, dados): 

        for i in range(len(dados)):

            dado = dados[i]
                
            celula = tabela.cell(dado["celula"][0], dado["celula"][1])
            paragrafo = celula.paragraphs[0]  

            Documento.criar_texto(paragrafo, dado["conteudo"], dado["negrito"], px = 10, fonte = "Cambria")

        return tabela
    
    @staticmethod
    def salvar_pdf(self):
        pdf = self.endereco.replace("WORD", "PDF")
        pdf = pdf.replace(".docx", ".pdf")
        #convert(docx, pdf)
        return pdf
                   
class Termo(Documento):
    
    def __init__(self, contratado, endereco, ordem, contrato, objeto, af, mensagem, tipo):
        super().__init__(contratado)
        self.ordem = ordem
        self.endereco = endereco
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.tipo = tipo


    def criar_arquivo(self, model):
        
        doc = Document(model)
        
        def definir_tabela(self):

            if self.tipo.lower() == "ata":
                campo = "ATA N°"
            else:
                campo = "CONTRATO N°" 

            dados = [{"celula": (1, 0), "conteudo": campo, "negrito": True}, 
                     {"celula": (1, 1), "conteudo": self.contrato, "negrito": False }, 
                     {"celula": (2, 1), "conteudo": self.contratado, "negrito": False }, 
                     {"celula": (3, 1), "conteudo": self.objeto, "negrito": False }, 
                     {"celula": (4, 1), "conteudo": self.af, "negrito": False}, 
                     {"celula": (6, 1), "conteudo": self.mensagem, "negrito": False }, 
                    ]
            
            doc.tables[0] = Documento.modificar_tabela(doc.tables[0], dados)
            
        def definir_data(self):
            data_texto = "BATAGUASSU/MS, " + Documento.encontrar_data_de_hoje_em_extenso()
            Documento.criar_texto(doc.add_paragraph(), data_texto, negrito = True, posicionamento = "Direita")
        
        def adicionar_espaco(self, quant):
            for i in range(quant):
                doc.add_paragraph("")
        
        def definir_gestor(self):
            print(self.tipo)
            if self.tipo.lower() == "contrato":
                Documento.adicionar_linha_de_assinatura(doc.add_paragraph(), "RONALDO DE SOUZA MARCÍLIO\nGESTOR DE CONTRATOS")
            else:
                Documento.adicionar_linha_de_assinatura(doc.add_paragraph(), "MURILO SOARES DE OLIVEIRA\nGESTOR DE ATAS")
                
        definir_tabela(self)
        definir_data(self)
        adicionar_espaco(self, 3)
        definir_gestor(self)
        
        return doc 
    
    def salvar_pdf(self):
        docx = self.endereco
        pdf = docx.replace("WORD", "PDF")
        pdf = pdf.replace(".docx", ".pdf")
        print(f"Convertendo termo para PDF.... *.✧*.✧.*.✧*.✧*.✧*.✧.,*.✧*. {self.contratado} - AF {self.af}")
        convert(docx, pdf)
        
class Relatorio(Documento):
    def __init__(self, contratado, liquidacao, valor, data):
        super().__init__(contratado)
        self.liquidacao = liquidacao
        self.valor = valor
        self.data = data 
