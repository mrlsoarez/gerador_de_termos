from docx import Document 

from Classes.Arquivo import Arquivo

class Documento():

    mapa = {
        "contratado": "B4",
        "n_contrato": "E4",
        "objeto": "F4",
        "numero_empenho": "A8",
        "numero_liquidacao": "A12",
        "data_liquidacao": "B12",
        "valor_bruto_liquidacao": "C12",
        "tipo_nota": "D16",
        "numero_af": "A20",
        "tipo": "B5"
    }

    def __init__(self, contratado, contrato, objeto, af, mensagem, gestor, liq, data, valor, tipo):
        # Info Termo
        self.contratado = contratado
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.gestor = gestor

        # Info Relatorio
        self.liq = liq 
        self.data = data 
        self.valor = valor

        self.tipo = tipo

class Termo(): 

    def __init__(self, termo, modelo, endereco):
        pass 

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

class Relatorio():

    def __init__(self, relatorio):
        pass   