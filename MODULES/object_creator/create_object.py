
from openpyxl import load_workbook
import win32api

from docx import Document
from docx2pdf import convert
from ENV.environment import pegar_modelos, atualizar_numero_protocolo
from MODULES.document_dealer.doc_dealer import DocHelper

import locale
import os
import shutil
"""

class Termo:


    def __init__(self, ordem, contrato, contratado, objeto, af, mensagem, gestor):
        self.ordem = ordem
        self.contratado = contratado 
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.gestor = gestor 

    def setRelatorioInfo(self, liquidacao, valor, data):
        self.liquidacao = liquidacao 
        self.valor = valor 
        self.data = data
 
    def checar_se_arquivo_existe(self):
        nome_arquivo = f'{self.ordem}. {self.contratado} - AF {self.af[:-3]}' + ".docx"
        if (os.path.isfile(nome_arquivo)): return True 

    def copiar_arquivo(self, arquivo, numero_protocolo = None, mesmo_protocolo = False):

        if (arquivo == 'termo'): 
            nome_arquivo = f'{self.ordem}. {self.contratado} - AF {self.af[:-3]}'
            endereco_copia = os.getcwd() + rf"\{nome_arquivo}" + ".docx"
           
        elif (arquivo == 'protocolo'): 
            endereco_copia = os.getcwd() + rf"\Protocolo N° {numero_protocolo} - Tesouraria.docx"
        
        def image_in_header(doc):
            section = doc.sections[0]
            header = section.header
            
            header.paragraphs[0].clear()

            paragraph = header.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  

            run = paragraph.add_run()
            run.add_picture(r"C:\Users\Usuario\Pictures\HEADER.png", width=Inches(6)) 
      
        def criar_tabela(doc):

        return endereco_copia
    
    def criar_termo(self, tipo, impressao = False):
        
        if (self.checar_se_arquivo_existe()): return 
        
        termo = self.copiar_arquivo("termo")

        doc = Document(termo)

        

        def adicionar_espaco():
            doc.add_paragraph("")
        
        def adicionar_data():
            data = DocHelper.encontrar_data_de_hoje_em_extenso()
            paragrafo = doc.add_paragraph()
            DocHelper.criar_texto(paragrafo, data, negrito=True, posicionamento = "Direita", fonte = "Cambria")
            pass 

        def adicionar_assinatura():
            paragrafo = doc.add_paragraph()
            DocHelper.adicionar_linha_de_assinatura(paragrafo, self.gestor)

        def imprimir_termo(termo, arquivo):
            pergunta = input(f"O termo --> {self.contratado} - AF {self.af} está pronto para ser impresso! Deseja imprimir? (Y/N).: ").lower()
            print(termo, arquivo)
            if (pergunta == "y"): 
                file_path = arquivo
            
                win32api.ShellExecute(
                    0,
                    "print",
                    file_path,
                    None,
                    ".",
                    0
                )
                
                print("Imprimindo... ", arquivo)
            
        definir_tabela()
        adicionar_espaco()
        adicionar_data()
  
        adicionar_espaco()
        adicionar_assinatura()

        def criar_assinatura(doc):
            assinatura = doc.add_paragraph("_________________________________________")
            assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
            gestor = doc.add_paragraph(self.gestor)
            gestor.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = gestor.runs[0]
            run.bold = True
        
        def footer(doc):
            section = doc.sections[0]
            footer = section.footer

    
            footer.paragraphs[0].clear()

    
            paragraph_img = footer.add_paragraph()
            paragraph_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = paragraph_img.add_run()
            run_img.add_picture(r"C:\Users\Usuario\Pictures\LINHA.png", width=Inches(6))

            paragraph_text = footer.add_paragraph()
            paragraph_text.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run_text = paragraph_text.add_run(
                "Avenida Aquidauana, Nº 1001 - Centro | Fone: (67) 3541-5100\n"
                "CEP 79.780-000 | CNPJ 03.576.220/0001-56\n"
                "www.bataguassu.ms.gov.br | gabinete@bataguassu.ms.gov.br"
            )

            # Optional formatting
            font = run_text.font
            font.size = Pt(8)
            font.name = "Arial"
        
        
            
        nome_arquivo = f"{Termo.index}. {self.contratado} - AF {self.af[:4]}"
        Termo.index += 1

        image_in_header(doc)
        criar_titulo(doc)
        doc.add_paragraph("")
        criar_tabela(doc)
        doc.add_paragraph("")
        criar_data(doc)
        doc.add_paragraph("") 
        criar_assinatura(doc)
        footer(doc)

        doc.save(nome_arquivo + ".docx")
        
    def salvar_pdf(docx):
        pdf = docx.replace("WORD", "PDF") + ".pdf"
        docx = docx + ".docx"
        convert(docx, pdf)
        return pdf         

    @staticmethod
    def criar_relatorio(termos, numero_protocolo, mesmo_protocolo):
        
        protocolo = Termo.copiar_arquivo(None, "protocolo", numero_protocolo, mesmo_protocolo)
        doc = Document(protocolo)

        def converter_currency(valor):
            locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")
            return locale.currency(float(valor), grouping =True)
        
        def adicionar_protocolo():
            substituir_protocolo = doc.paragraphs[1]
            substituir_protocolo.text = ""
            DocHelper.criar_texto(substituir_protocolo, f"PROTOCOLO DE RECEBIMENTO - NÚMERO {numero_protocolo}", negrito = True)
   
        def criar_tabela():
            
            tabela = doc.tables[0]
            for i in range(len(termos)):
                
                nova_linha = tabela.add_row()
            
                coluna_um = nova_linha.cells[0].paragraphs[0]
                coluna_dois = nova_linha.cells[1].paragraphs[0]
                coluna_tres = nova_linha.cells[2].paragraphs[0]
                coluna_quatro = nova_linha.cells[3].paragraphs[0]
                
                DocHelper.criar_texto(coluna_um, termos[i].contratado,  px = 8, negrito = True, fonte = "Arial")
                DocHelper.criar_texto(coluna_dois, termos[i].liquidacao,  px = 8, negrito = True, fonte = "Arial")
                DocHelper.criar_texto(coluna_tres, termos[i].data,  px = 8, negrito = True, fonte = "Arial")
                DocHelper.criar_texto(coluna_quatro, converter_currency(termos[i].valor),  px = 8, negrito = True, fonte = "Arial")
                
           
        adicionar_protocolo() 
        criar_tabela()

        ask_it = str(input("Atualizar o protocolo? (Y/N) --> ")).lower()

        if ask_it == "y": atualizar_numero_protocolo()
        doc.save(protocolo)

        print()

def capturar_info_planilha(localizacao_planilha, numero_especifico = False):

    def formatar_data(objeto):
        return objeto.date().strftime("%d/%m/%Y")
    
    def customizar_mensagem(tipo_nota):
        if (tipo_nota == "Locação"): return "Por este instrumento, em caráter DEFINITIVO, atestamos que a locação acima identificada atende às exigências contratuais."
        return f"Por este instrumento, em caráter DEFINITIVO, atestamos que os {tipo_nota.lower()} acima identificados atendem às exigências contratuais."
    
    PLANILHA = load_workbook(localizacao_planilha, data_only = True)
    TERMOS = []

    for nome in PLANILHA.sheetnames: 
        if nome != "Base" and nome != "Fiscais":

            if numero_especifico != False: 
                try: 
                    SHEET = PLANILHA[numero_especifico]
                except: 
                    pass  
                numero_especifico = str(int(numero_especifico) + 1)
            else: 
                SHEET = PLANILHA[nome]

            if (SHEET["A4"].value == None or SHEET ["A30"].value == "Sim"): 
                continue 
 
            ordem = nome
            contrato = SHEET["E4"].value 
            contratado = SHEET["B4"].value
            mensagem = customizar_mensagem(str(SHEET["D16"].value))
            objeto = SHEET["F4"].value 
            af = SHEET["A20"].value

            liquidacao = SHEET["A8"].value 
            data_liquidacao = formatar_data(SHEET["B12"].value)
            valor = SHEET["C12"].value

            if ("ATA" in localizacao_planilha):
                gestor = "MURILO SOARES DE OLIVEIRA\nGESTOR DE ATA"
            else: 
                gestor = "RONALDO DE SOUZA MARCILIO\nSETOR DE CONTRATOS"
        
            novo_termo = Termo(ordem, contrato, contratado, objeto, af, mensagem, gestor)
            novo_termo.setRelatorioInfo(liquidacao, valor, data_liquidacao)

            TERMOS.append(novo_termo)

    return TERMOS
"""