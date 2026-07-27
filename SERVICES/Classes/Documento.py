
from openpyxl import load_workbook

from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

from ENV.environment import pegar_modelos

import shutil
import locale

from docx import Document
from docx2pdf import convert
from copy import deepcopy

from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Pt


class Documento: 
    
    def __init__(self, contratado):
        self.contratado = contratado
        
    def set_af(self, dado):
        self.af = dado
    
    def set_arq_nome(self, dado):
        self.arq = dado 
        
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
    
    def __init__(self, contratado, endereco, ordem, contrato, objeto, af, mensagem, gestor, tipo, modelo):
        super().__init__(contratado)
        self.ordem = ordem
        self.endereco = endereco
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.gestor = gestor
        self.tipo = tipo
        self.modelo = modelo 

    def criar_arquivo(self):
        
        doc = Document(self.modelo)
        
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
            Documento.adicionar_linha_de_assinatura(doc.add_paragraph(), self.gestor)
            
                
        definir_tabela(self)
        definir_data(self)
        adicionar_espaco(self, 3)
        definir_gestor(self)
        
        doc.save(self.endereco)
    
    def salvar_pdf(self):
        docx = self.endereco
        pdf = docx.replace("WORD", "PDF")
        pdf = pdf.replace(".docx", ".pdf")
        print(f"Convertendo termo para PDF.... *.✧*.✧.*.✧*.✧*.✧*.✧.,*.✧*. {self.contratado} - AF {self.af}")
        convert(docx, pdf)
        
class Relatorio(Documento):
    def __init__(self, contratado, liquidacao, valor, data, modelo, protocolo, endereco):
        super().__init__(contratado)
        self.liquidacao = liquidacao
        self.valor = valor
        self.data = data 
        self.modelo = modelo
        self.protocolo = protocolo
        self.endereco = endereco
    
    def criar_arquivo(self, doc, termos):
        
        def converter_currency(self, valor):
            locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")
            return locale.currency(float(valor), grouping =True)
                    
        def adicionar_protocolo(self):
            substituir_protocolo = doc.paragraphs[1]
            substituir_protocolo.text = ""
            Documento.criar_texto(substituir_protocolo, f"PROTOCOLO DE RECEBIMENTO - NÚMERO {self.protocolo}", negrito = True)
               
        def criar_tabela(self, termos):
            
            def limpar_tabela(tabela):
                while len(tabela.rows) > 1:
                    tr = tabela.rows[1]._tr
                    tr.getparent().remove(tr)
                    
            limpar_tabela(doc.tables[0])
            
            for i in range(len(termos)):
                
                tabela = doc.tables[0]   
                nova_linha = tabela.add_row()
                                
                coluna_um = nova_linha.cells[0].paragraphs[0]
                coluna_dois = nova_linha.cells[1].paragraphs[0]
                coluna_tres = nova_linha.cells[2].paragraphs[0]
                coluna_quatro = nova_linha.cells[3].paragraphs[0]
                                    
                Documento.criar_texto(coluna_um, termos[i].contratado,  px = 8, negrito = True, fonte = "Arial")
                Documento.criar_texto(coluna_dois, termos[i].liquidacao,  px = 8, negrito = True, fonte = "Arial")
                Documento.criar_texto(coluna_tres, termos[i].data,  px = 8, negrito = True, fonte = "Arial")
                Documento.criar_texto(coluna_quatro, converter_currency(self, termos[i].valor),  px = 8, negrito = True, fonte = "Arial")
                                   
    
        adicionar_protocolo(self) 
        criar_tabela(self, termos)
        
        doc.save(self.endereco)