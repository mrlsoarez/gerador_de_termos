
MAPEAMENTO = {
    "contratado": "B2",
    "n_contrato": "E2",
    "objeto": "F2",
    "numero_empenho": "A6",
    "numero_liquidacao": "A10",
    "data_liquidacao": "B10",
    "valor_bruto_liquidacao": "C10",
    "tipo_nota": "D14",
    "numero_af": "A18",
}

class CriadorDeArquivo:
    
    def __init__(self, ordem, contrato, contratado, objeto, af, mensagem, liquidacao, valor, data, gestor):
        self.ordem = ordem
        self.contratado = contratado 
        self.contrato = contrato 
        self.objeto = objeto 
        self.af = af
        self.mensagem = mensagem
        self.liquidacao = liquidacao 
        self.valor = valor 
        self.data = data
        self.gestor = gestor 

def verificar_se_arquivo_existe():
    pass