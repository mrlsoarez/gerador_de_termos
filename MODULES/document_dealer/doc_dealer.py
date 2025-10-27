from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Pt

class DocHelper:

    def mudar_fonte(texto, name_font):
        texto.font.name = name_font
    
    def mudar_tamanho(texto, px): 
        texto.font.size = Pt(px)

    def criar_texto(paragrafo, texto, negrito = False, posicionamento = None, px = None, fonte = None):

        def deixar_negrito(run): 
           run.bold = True 

        def alinhar_texto(paragrafo, alinhado):
            if (alinhado == "Centro"): 
                paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif (alinhado == "Direita"):
                paragrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT 

        run = paragrafo.add_run(texto)

        if (negrito): deixar_negrito(run)
        if (posicionamento != None): alinhar_texto(paragrafo, posicionamento)
        if (px != None): DocHelper.mudar_tamanho(run, px)
        if (fonte != None): DocHelper.mudar_fonte(run, fonte)

    def encontrar_data_de_hoje_em_extenso():
        dicionario = {
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
                
        data = str(date.today())

        dia = data[0] + data[1]
        ano = data[:4]
        mes = dicionario[f"{data[5]}{data[6]}"]

        return f"BATAGUASSU/MS, {dia} de {mes} de {ano}"

    def adicionar_linha_de_assinatura(paragrafo, assinador): 
        DocHelper.criar_texto(paragrafo, f"_____________________________\n{assinador}", negrito = True, posicionamento = "Centro", fonte = "Cambria")
 

    def modificar_tabela(tabela, dados): 

        for i in range(len(dados)):

            dado = dados[i]
                
            celula = tabela.cell(dado["celula"][0], dado["celula"][1])
            paragrafo = celula.paragraphs[0]  

            DocHelper.criar_texto(paragrafo, dado["conteudo"], dado["negrito"], px = 10, fonte = "Cambria")

        return tabela
         