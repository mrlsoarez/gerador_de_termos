from Classes.Documento import Documento 

# Lê a planilha e verifica se existem campos vazios     

def analisarPlanilha(inicializador):

    def seNaoExistemCamposVazios(mapa, planilha, sheet):
        verificacao = True 

        if (planilha[sheet]["A4"].value == None or (planilha[sheet]["B4"].value == None)):
            verificacao = False 
            return verificacao

        for keys in mapa:
            if (planilha[sheet][mapa[keys]].value == None) and keys != "contratado":
                print(f"O campo '{keys}' está vazio na planilha {sheet}")
                verificacao = False  
                break 

        return verificacao 
    
    DADOS = {"contrato": [], "ata": []}

    planilha = inicializador.planilha
    mapa = Documento.mapa 

    for sheet in planilha.sheetnames:
        try: 
            parse = int(sheet)
        except: 
            pass 
        else: 
            if (seNaoExistemCamposVazios(mapa, planilha, sheet)):
                doc = Documento(
                    ordem = sheet,
                    contratado = planilha[sheet][mapa["contratado"]].value,
                    contrato = planilha[sheet][mapa["n_contrato"]].value,
                    objeto = planilha[sheet][mapa["objeto"]].value,
                    af = planilha[sheet][mapa["numero_af"]].value,
                    tipo_nota = planilha[sheet][mapa["tipo_nota"]].value,
                    liq = planilha[sheet][mapa["numero_liquidacao"]].value,
                    data = planilha[sheet][mapa["data_liquidacao"]].value,
                    valor = planilha[sheet][mapa["valor_bruto_liquidacao"]].value,
                    tipo = planilha[sheet][mapa["tipo"]].value
                )
                
                
                DADOS[doc.tipo.lower()].append(doc)

    return DADOS

