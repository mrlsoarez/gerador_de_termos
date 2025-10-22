from openpyxl import load_workbook

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Inches
from docx import Document
from docx2pdf import convert

import locale
import os

class Termo:

    index = 1 

    def __init__(self, contrato, contratado, objeto, af, mensagem, gestor):
        self.contratado = contratado 
        self.contrato = contrato 
        self.mensagem = mensagem
        self.objeto = objeto 
        self.af = af
        self.gestor = gestor 

    def setRelatorioInfo(self, liquidacao, valor, data):
        self.contratado = self.contratado
        self.liquidacao = liquidacao 
        self.valor = valor 
        self.data = data 

    def criar_doc_termo(self, tipo):

        doc = Document()
        
        def image_in_header(doc):
            section = doc.sections[0]
            header = section.header
            
            header.paragraphs[0].clear()

            paragraph = header.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  

            run = paragraph.add_run()
            run.add_picture(r"C:\Users\Usuario\Pictures\HEADER.png", width=Inches(6)) 
      
        def criar_tabela(doc):

            tabela = doc.add_table(rows=6, cols=2)
            tabela.style = "Table Grid"
            celula = tabela.cell(0, 0).merge(tabela.cell(0, 1))
            p = celula.paragraphs[0]
            run = p.add_run("1- IDENTIFICAÇÃO")
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if "credenciamento" in self.objeto.lower():
                campo = "CREDENCIAMENTO N°"
            elif tipo == "ata":
                campo = "ATA N°"
            else:
                campo = "CONTRATO N°" 

            campos = [campo, "CONTRATADO", "OBJETO", "AUTORIZAÇÃO DE FORNECIMENTO"]
            for i, campo in enumerate(campos, start=1):
                cell = tabela.cell(i, 0)
                p = cell.paragraphs[0]
                run = p.add_run(campo)
                run.bold = True

            tabela.cell(1, 1).text = self.contrato
            tabela.cell(2, 1).text = self.contratado 
            tabela.cell(3, 1).text = self.objeto
            tabela.cell(4, 1).text = self.af

            cell = tabela.cell(5, 0).merge(tabela.cell(5, 1))
            p = cell.paragraphs[0]
            run = p.add_run("2- CUMPRIMENTO DAS OBRIGAÇÕES")
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # --- Linha 7: primeiro texto dentro da tabela (mesclar colunas) ---
            cell = tabela.add_row().cells[0].merge(tabela.add_row().cells[1])
            p = cell.paragraphs[0]
            p.add_run(self.mensagem)

            # --- Linha 8: segundo texto dentro da tabela (mesclar colunas) ---
            cell = tabela.add_row().cells[0].merge(tabela.add_row().cells[1])
            cell.text = (
                "Constitui ainda eficácia liberatória de todas as obrigações estabelecidas em "
                "contratado referentes ao objeto acima mencionado, exceto as garantias legais, "
                "bem como autorizamos a restituição de todas as garantias e/ou caução prestadas."
            )
            
        def criar_titulo(doc):
          titulo = doc.add_paragraph("TERMO DE RECEBIMENTO DEFINITIVO")
          titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
          run = titulo.runs[0]
          run.bold = True
          run.font.size = Pt(14)
        
        def criar_data(doc):

            def truncate_date():

                today = str(date.today())
                
                def get_month(today):
                    month = today[5] + today[6]
                    months = {
                    "01": "JANEIRO",
                    "02": "FEVEREIRO",
                    "03": "MARÇO",
                    "04": "ABRIL",
                    "05": "MAIO",
                    "06": "JUNHO",
                    "07": "JULHO",
                    "08": "AGOSTO",
                    "09": "SETEMBRO",
                    "10": "OUTUBRO",
                    "11": "NOVEMBRO",
                    "12": "DEZEMBRO"
                    }
                    for keys in months:
                        if keys == month:
                            return months[keys]

                month = get_month(today)        
                string = "BATAGUASSU/MS, "
                string += today[8:]  
                string += f" de {month} de 2025."
                return string
            
            text = truncate_date()
            data = doc.add_paragraph(text)
            data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            truncate_date()

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
        
        arquivo_docx = os.path.abspath(f"{nome_arquivo}")
        Termo.salvar_pdf(arquivo_docx)

    
    def salvar_pdf(docx):
        pdf = docx.replace("WORD", "PDF") + ".pdf"
        docx = docx + ".docx"
        convert(docx, pdf)


    @staticmethod
    def gerar_relatorio(termos, numero_protocolo):

             # Update número de protocolo
      def update_protocol():
            fonte = r"C:\Program Files (x86)\numero_protocolo\protocolo.txt"
            numero_atualizado = int(numero_protocolo) + 1
            with open(fonte, "r+") as txt:
                txt.write(str(numero_atualizado))
       
      # Deixar elemento em negrito
      def make_it_bold(p):
           run = p.runs[0]
           run.bold = True 
           return run 
      
      # Mudar tamanho de fonte
      def change_font_size(p, px):
          run = p.runs[0]
          run.font.size = Pt(px)

      #############################################################################
      # -> Criando DOC

      doc = Document()

      def criar_cabeçalho(doc):

        def image_in_header():
            section = doc.sections[0]
            header = section.header
            run = header.paragraphs[0].add_run()
            run.add_picture(r"C:\Users\Usuario\Pictures\HEADER.png")

        image_in_header()

        titulo_protocolo = doc.add_paragraph("SETOR DE CONTRATOS")
        titulo_protocolo_numero = doc.add_paragraph(f"PROTOCOLO DE RECEBIMENTO - NUMERO {numero_protocolo}")
      
        make_it_bold(titulo_protocolo)
        make_it_bold(titulo_protocolo_numero)
      
        doc.add_paragraph("__________________________________________________________________________________________")

        mensagem_tesouraria_um = make_it_bold(doc.add_paragraph("Ao setor de Tesouraria, "))
        mensagem_tesouraria_dois = doc.add_paragraph(f"Encaminhamos, por meio deste, as notas fiscais em anexo para as devidas providências quanto à análise e efetivação do pagamento. ")
        mensagem_tesouraria_tres = make_it_bold(doc.add_paragraph("Relação dos documentos: "))

      def criar_tabela(doc):

        def converter_currency(valor):
            locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")
            return locale.currency(float(valor), grouping =True)

        table = doc.add_table(rows = len(termos) + 1, cols = 4)
        table.style = "Table Grid"
        table.autofit = "False"

        # TABELA COLUNA PADRÃO
        campos = ["EMPRESA", "N° DE LIQUIDAÇÃO", "DATA", "VALOR"]
        for i in range(0, 4):
           cell = table.cell(0, i)  
           if i == 0: cell.width = Inches(5)
           p = cell.paragraphs[0]
           run = p.add_run(campos[i])
           run.bold = True
           change_font_size(cell.paragraphs[0], 9)

        #TABELA COLUNA DINÂMICA
       
        for i in range(0, len(termos)):
          for q in range(0, 4):  
            if q == 0: table.cell(i+1, q).text = termos[i].contratado
            elif q == 1: table.cell(i+1, q).text = termos[i].liquidacao
            elif q == 2: table.cell(i+1, q).text = termos[i].data
            elif q == 3: table.cell(i+1, q).text = converter_currency(termos[i].valor)
            cell = table.cell(i+1, q) 
            change_font_size(cell.paragraphs[0], 9)
          
      def adicionar_paragrafo(doc):
        doc.add_paragraph("")
    
      def mensagem_tesouraria(doc):
        mensagem_tesouraria_quatro = make_it_bold(doc.add_paragraph("Atenciosamente, "))
      
      def criar_assinaturas(doc):
        assinatura_murilo = doc.add_paragraph("_________________________________________")
        assinatura_murilo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        murilo = doc.add_paragraph("MURILO SOARES DE OLIVEIRA \n SETOR DE CONTRATOS")
        murilo.alignment = WD_ALIGN_PARAGRAPH.CENTER
     
        doc.add_paragraph("")
        doc.add_paragraph("DATA DE ASSINATURA: Bataguassu/MS ____ do _____ de 2025")
        doc.add_paragraph("")

        assinatura_tesouraria = doc.add_paragraph("_________________________________________")
        assinatura_tesouraria.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tesouraria = doc.add_paragraph("ASSINATURA DE RECEBIMENTO \n SETOR DE TESOURARIA")
        tesouraria.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        make_it_bold(murilo)
        make_it_bold(tesouraria)

      criar_cabeçalho(doc)
      criar_tabela(doc)
      adicionar_paragrafo(doc)
      mensagem_tesouraria(doc)
      criar_assinaturas(doc)
      
      nome_arquivo = f"Protocolo N° {numero_protocolo} - Tesouraria"
      doc.save(f"{nome_arquivo}.docx")
      
      ask_it = str(input("Atualizar o protocolo? (Y/N) --> ")).lower()

      if ask_it == "y": update_protocol()

      
    
def capturar_info_planilha(localizacao_planilha):

    def formatar_data(objeto):
        return objeto.date().strftime("%d/%m/%Y")
    
    def customizar_mensagem(tipo_nota):
        if (tipo_nota == "Locação"): return "Por este instrumento, em caráter DEFINITIVO, atestamos que a locação acima identificada atende às exigências contratuais."
        return f"Por este instrumento, em caráter DEFINITIVO, atestamos que os {tipo_nota.lower()} acima identificados atendem às exigências contratuais."
    
    PLANILHA = load_workbook(localizacao_planilha, data_only = True)
    TERMOS = []

    for nome in PLANILHA.sheetnames: 
        if nome != "Base" and nome != "Fiscais":

            SHEET = PLANILHA[nome]

            if (SHEET["A4"].value == None): 
                continue 
             
            contrato = SHEET["E4"].value 
            contratado = SHEET["B4"].value
            mensagem = customizar_mensagem(str(SHEET["D16"].value))
            objeto = SHEET["F4"].value 
            af = SHEET["A20"].value

            liquidacao = SHEET["A8"].value 
            data_liquidacao = formatar_data(SHEET["B12"].value)
            valor = SHEET["C12"].value

            if ("ATA" in localizacao_planilha):
                gestor = "GESTOR DE ATA\nMURILO SOARES DE OLIVEIRA"
            else: 
                gestor = "GESTOR DE CONTRATO\nRONALDO DE SOUZA MARCILIO"
        
            novo_termo = Termo(contrato, contratado, objeto, af, mensagem, gestor)
            novo_termo.setRelatorioInfo(liquidacao, valor, data_liquidacao)

            TERMOS.append(novo_termo)

    return TERMOS