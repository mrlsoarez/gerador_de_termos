from MODULES.GerenciarArquivos import GerenciarArquivos

class Verificador:
    def __init__(self, contratado, sheet):
        
        self.contratado = contratado
        self.sheet = sheet
        
        self.mapa = {
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
        """
        
        self.mapa.portaria = {
            "fornecedor": "B",
            "n_contrato": "C",
            "objeto": "D",
            "portaria": "E",
            "secretaria": "F",
            "fiscal_principal": "G",
            "fiscal_suplente": "H",
        }
        """

    def set_info(self):
        self.contratado = self.sheet[self.mapa['contratado']].value
        self.af = self.sheet[self.mapa['numero_af']].value
        self.arq = f"{self.contratado.replace("/", "")} - AF {self.af[:4]}.docx"
        pass
    
    def checar_campos_planilha(self):
        for campo, celula in self.mapa.items():
            valor = self.sheet[celula].value 
            if (valor == None):
                return {'resultado': False, 'mensagem': f'{self.sheet} -> O campo "{campo}" desta planilha está vazio e sem conteúdo. Necessário preencher e tentar novamente!' }
        return {'resultado': True, 'mensagem': ""}
    
   