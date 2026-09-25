from Classes.Documento import Documento 
import win32com.client

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
                
                try:
                    DADOS[doc.tipo.lower()].append(doc)
                except: 
                    pass

    return DADOS

def resetarPlanilha(ordens, inicializador): 
    
    caminho = inicializador.caminhoPlanilha

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try: 
        wb = excel.Workbooks.Open(caminho)
        regularPlanilha(wb)
    except Exception as e: 
        print(e)
    else: 
        wb.Save()
        wb.Close()
    finally: 
        excel.quit()
    """
    
    try:

        wb = excel.Workbooks.Open(caminho)
        for key in ordens:
            for ordem in ordens[key]:
                try:
                    wb.Worksheets(str(ordem)).Delete()
                except Exception as e:
                    print(f"Erro ao remover '{ordem}': {e}")
        
        numero = 1
        quantidade = wb.Worksheets.Count
        for i in range(1, quantidade + 1):
            sheet = wb.Worksheets(i)
            if "base" not in sheet.Name.lower():
                novo_nome = str(numero)
                sheet.Name = novo_nome
                numero += 1
        
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()
        
    """
def regularPlanilha(planilha):
    
    def encontrarUltimaOrdem(array):
        array.sort(key=int)
        ultimoIndex = len(array) - 1
        return int(array[ultimoIndex])
    
    sheetsProntas = []
    modelo = planilha.Sheets("Modelo")
    print(modelo)
    for sheet in planilha.sheetnames:
        try: 
            parse = int(sheet)
        except: 
            pass 
        else: 
            sheetsProntas.append(sheet)
    
    ultimaOrdem = encontrarUltimaOrdem(sheetsProntas)
    
    for i in range(ultimaOrdem+1, 11):
        modelo.Copy(After=planilha.Sheets(planilha.Sheets.Count))
        planilha.Sheets(planilha.Sheets.Count).Name = str(i)
    
   
    
    #planilha.save(rf"{inicializador.pastaPlanilha}\a.xlsx")
