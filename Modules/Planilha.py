from Classes.Documento import Documento 
import win32com.client

# Lê a planilha e verifica se existem campos vazios     

def analisarPlanilha(inicializador):

    def seNaoExistemCamposVazios(mapa, planilha, sheet):
        verificacao = True 
        if (planilha[sheet][mapa["processo"]].value == None or (planilha[sheet][mapa["contratado"]] == None)):
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
    
    mapa = Documento.mapa 
    
    def removerInformacoes(wb, ordens): 
        for key in ordens:
            for ordem in ordens[key]:
                try:
                    wb.Worksheets(str(ordem)).Delete()
                except Exception as e:
                    print(f"Erro ao remover '{ordem}': {e}")
        for s in wb.Worksheets: 
            print(s.Range(mapa["processo"]))
            """
             
            numero = 1
            quantidade = wb.Worksheets.Count
            for i in range(1, quantidade + 1):
                sheet = wb.Worksheets(i)
                if "base" not in sheet.Name.lower():
                    novo_nome = str(numero)
                    sheet.Name = novo_nome
                    numero += 1
            """
    def regularInformacoes(wb, total=10):
        modelo = wb.Sheets("Modelo")
        for s in wb.Sheets:
            print()
        """
        
        if wb.ProtectStructure:
            raise RuntimeError("Workbook structure is protected.")


        # 1) Collect numbered sheets, sorted by their current number
        numeradas = []
        for s in wb.Sheets: 
            print(s, wb.Sheets(s))
            try: 
                parse = int(wb.Sheets(s).Name) 
            except: 
                pass 
            else: 
                numeradas.append(wb.Sheets(s).Name)
        #numeradas = [s for s in wb.Sheets if s.Name.isdigit()]
        #numeradas.sort(key=lambda s: int(s.Name))
        #print("Original order:", [s.Name for s in numeradas])
        print(numeradas)

        # 2) Rename to temporary names first (avoids collisions like 8 -> 4 when a "4" exists)
        for k, s in enumerate(numeradas):
            s.Name = f"__tmp_{k}"

        # 3) Final names 1..N, and reorder the tabs
        for k, s in enumerate(numeradas, start=1):
            s.Name = str(k)
        for prev, s in zip(numeradas, numeradas[1:]):
            s.Move(After=prev)

        ultima = numeradas[-1] if numeradas else modelo
        print("Preserved sheets:", [s.Name for s in numeradas])

        # 4) Add empty copies of Modelo up to `total`
        for i in range(len(numeradas) + 1, total + 1):
            modelo.Copy(After=ultima)
            nova = wb.ActiveSheet          # the copy becomes the active sheet
            nova.Name = str(i)
            ultima = nova                  # next copy goes after this one
            print(f"Created sheet {i}")

        print("Final order:", [s.Name for s in wb.Sheets])
        """
    
        
        #planilha.save(rf"{inicializador.pastaPlanilha}\a.xlsx")
    caminho = inicializador.caminhoPlanilha

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    
    try: 
        wb = excel.Workbooks.Open(caminho)
        removerInformacoes(wb, ordens)
        regularInformacoes(wb)
    except Exception as e: 
        print(e)
    else: 
        wb.Save()
        wb.Close()
    finally: 
        excel.Quit()
    

    """
    
    try:

       
        
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()
    """
        
  

