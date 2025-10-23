from openpyxl import load_workbook

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Inches
from docx import Document
from docx2pdf import convert

import locale
import os


"""
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
        self.liquidacao = liquidacao 
        self.valor = valor 
        self.data = data 

    def criar_doc_termo(self, tipo):

        doc = Document()
        
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

        nome_arquivo = f"{Termo.index}. {self.contratado} - AF {self.af[:4]}"
        Termo.index += 1

        criar_titulo(doc)
        doc.add_paragraph("")
        criar_tabela(doc)
        doc.add_paragraph("")
        criar_data(doc)
        doc.add_paragraph("") 
        criar_assinatura(doc)

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

class DocCreator: 
    
    index = 1
    
    def __init__(self, tabela_info):
        self.tabela_info = tabela_info
      
    def alinhar_elemento(elemento, alinhado):
        if (alinhado == "Center"): 
            elemento.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    
    def deixar_negrito(elemento): 
        elemento.bold = True
        return elemento 
    
    def criar_uma_tabela(self, doc):
        
        if (self.incremental):
            tabela = doc.add_table(rows= self.tabela_info["linhas"], cols = self.tabela_info["colunas"])
            tabela.style = "Table Grid"
        
        DETALHES_CAMPO = self.tabela_info["campos"]

        for detalhes in DETALHES_CAMPO:
            conteudo = DETALHES_CAMPO[detalhes]["conteudo"]
            celula_referenciada = DETALHES_CAMPO[detalhes]["celula_referenciada"]
            
            if (DETALHES_CAMPO[detalhes]["merge"]):
                range_do_merge = DETALHES_CAMPO[detalhes]["range_para_merge"]
                celula = tabela.cell(celula_referenciada[0], celula_referenciada[1]).merge(tabela.cell(range_do_merge[0], range_do_merge[1]))
            else: 
                celula = tabela.cell(celula_referenciada[0], celula_referenciada[1])
                
            paragraph = celula.paragraphs[0]
           
            if (DETALHES_CAMPO[detalhes]["negrito"]): paragraph.add_run(conteudo).bold = True 
            else: paragraph.add_run(conteudo)
            if (DETALHES_CAMPO[detalhes]["alinhado"]): DocCreator.alinhar_elemento(paragraph, DETALHES_CAMPO[detalhes]["alinhado"])
        
        range_inicio = self.tabela_info["range_dados_inicio"][0] 
        range_final = self.tabela_info["range_dados_final"][1] + 1
        dados = self.tabela_info["dados"]
        
        for i in range(range_inicio, range_final): 
            tabela.cell(i, 1).text = dados[i - 1]
           
        
        
                
     
        if (self.tabela_info["merge"] == True): 
            DETALHES_MERGE = self.tabela_info["detalhes_merge"] 
            for merge in DETALHES_MERGE:
                
                celula_referenciada = DETALHES_MERGE[merge]["celula_referenciada"]
                range_do_merge = DETALHES_MERGE[merge]["range_para_merge"]
                
                celula = tabela.cell(celula_referenciada[0], celula_referenciada[1]).merge(tabela.cell(range_do_merge[0], range_do_merge[1]))
                text = DETALHES_MERGE[merge]["conteudo"]
                
                paragraph = celula.paragraphs[0]
                
                if (DETALHES_MERGE[merge]["negrito"]): 
                    paragraph.add_run(text).bold = True 
                if (DETALHES_MERGE[merge]["alinhado"] != None):
                    DocCreator.alinhar_elemento(paragraph, DETALHES_MERGE[merge]["alinhado"])
                else: 
                    paragraph.add_run(text)
  
                    
        
                
            
                #celula.text = paragraph
                #if (DETALHES_MERGE[merge]["alinhado"] != ""): DocCreator.alinhar_elemento(celula, )
              

        #p = celula.paragraphs[0]
        #run = p.add_run("1- IDENTIFICAÇÃO")
        #run.bold = True
        #p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return tabela 
        

def objeto_teste():
    TERMOS = [
    Termo("CONTRATO 001/2025", "EMPRESA ALFA LTDA", "Fornecimento de material de escritório", "AF-1234", "Entrega concluída com sucesso.", "GESTOR DE CONTRATO\nRONALDO DE SOUZA MARCILIO"),
    Termo("ATA 002/2025", "SUPRIMENTOS BETA ME", "Registro de preços para gêneros alimentícios", "AF-2234", "Execução conforme cronograma.", "GESTOR DE ATA\nMURILO SOARES DE OLIVEIRA"),
    Termo("CONTRATO 003/2025", "SERVIÇOS GAMA EIRELI", "Prestação de serviços de limpeza", "AF-3234", "Serviços executados parcialmente.", "GESTOR DE CONTRATO\nRONALDO DE SOUZA MARCILIO"),
    Termo("ATA 004/2025", "ALIMENTOS DELTA LTDA", "Registro de preços para merenda escolar", "AF-4234", "Fornecimento regular e satisfatório.", "GESTOR DE ATA\nMURILO SOARES DE OLIVEIRA"),
    Termo("CONTRATO 005/2025", "TECNOLOGIA ÔMEGA S/A", "Licenciamento de software e suporte técnico", "AF-5234", "Sistema implantado e testado.", "GESTOR DE CONTRATO\nRONALDO DE SOUZA MARCILIO")
    ]

# Exemplo opcional: atribuir informações de relatório a cada termo
    TERMOS[0].setRelatorioInfo("LIQ-1001", 12500.00, "10/03/2025")
    TERMOS[1].setRelatorioInfo("LIQ-1002", 8500.50, "18/03/2025")
    TERMOS[2].setRelatorioInfo("LIQ-1003", 22000.00, "25/03/2025")
    TERMOS[3].setRelatorioInfo("LIQ-1004", 14200.75, "29/03/2025")
    TERMOS[4].setRelatorioInfo("LIQ-1005", 18750.00, "02/04/2025")
    
    return TERMOS 
def teste_doc():
    
    doc = Document()
    
    def get_tabela_info(obj, *args):
        return {
        "linhas": 8,
        "colunas": 2,
        "dados": obj,
        "incremental": False,
        "range_dados_inicio": (1,1),
        "range_dados_final": (1,4),
        "merge": True,
        "campos": { "primeiro_campo": {
                            "celula_referenciada": (0, 0),
                            "range_para_merge": (0, 1),
                            "merge": True,
                            "conteudo": "1 - IDENTIFICAÇÃO",
                            "negrito": True,
                            "alinhado": "Center",
                        },
                       "segundo_campo": {
                            "celula_referenciada": (1, 0),
                            "conteudo": "CONTRATO N°",
                            "merge": False,
                            "negrito": True,
                            "alinhado": False,
                        },
                        "terceiro_campo": {
                            "celula_referenciada": (2, 0),
                            "conteudo": "CONTRATADO",
                            "merge": False,
                            "negrito": True,
                            "alinhado": False,
                        },
                        "quarto_campo": {
                            "celula_referenciada": (3, 0),
                            "conteudo": "OBJETO",
                            "negrito": True,
                            "merge": False,
                            "alinhado": False,
                        }, 
                        "quinto_campo": {
                            "celula_referenciada": (4, 0),
                            "conteudo": "AUTORIZACAO DE FORNECIMENTO",
                            "merge": False,
                            "negrito": True,
                            "alinhado": False,
                        }, 
                        "sexto_campo": {
                            "celula_referenciada": (5, 0),
                            "range_para_merge": (5, 1),
                            "merge": True,
                            "conteudo": "2 - CUMPRIMENTO DAS OBRIGAÇÕES",
                            "negrito": True, 
                            "alinhado": "Center"
                        },
                        "setimo_campo": {
                            "celula_referenciada": (6, 0),
                            "range_para_merge": (6, 1),
                            "merge": True,
                            "conteudo": obj.mensagem, 
                            "negrito": False, 
                            "alinhado": False
                        },
                        "oitavo_campo": {
                            "celula_referenciada": (7, 0),
                            "range_para_merge": (7, 1),
                            "merge": True,
                            "conteudo": "Constitui ainda eficácia liberatória de todas as obrigações estabelecidas em contratado referentes ao objeto acima mencionado, exceto as garantias legais, bem como autorizamos a restituição de todas as garantias e/ou caução prestadas.",
                            "negrito": False, 
                            "alinhado": False
                        }
        }
    }
    
    termos = objeto_teste()
    
    for i in range(len(termos)):
        tabela_info = get_tabela_info(termos)
        doc = Document()
        for key in tabela_info:
            test = DocCreator(tabela_info)
        test.criar_uma_tabela(doc)

    

#teste_doc()

"""
"""
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
"""