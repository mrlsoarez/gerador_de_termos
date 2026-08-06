# Pasta onde ficará armazenado os termos, separados por datas

ROOT = r"C:\Users\Usuario\Documents\MRL"
TXT_PROTOCOLO = rf"{ROOT}\MODELOS BASE\2. CONTROLE DE PROTOCOLO - RELATÓRIO\protocolo.txt"

# Onde será guardado o documento pronto
def pegar_endereco_base(op = None):
    if (op == "4"):
        return rf"{ROOT}\2. DOCS - EQUIPE DE APOIO\2. PORTARIAS"
    
    return rf"{ROOT}\1. ANÁLISE DE PAGAMENTOS\2026"
  
def pegar_planilha_termo(arquivo):
    return rf"{ROOT}\MODELOS BASE\1. PLANILHAS - ANÁLISE FISCAL - ATA & CONTRATOS\{arquivo}"
    
def pegar_tipo_termo(prompt):
    print(prompt)
    if (prompt[0] == "1"):
        return [{"tipo": "CONTRATO", "arquivo": "ANÁLISE FISCAL.xlsx"}]
    if (prompt[0] == "2"): 
        return [{"tipo": "ATA", "arquivo": "ANÁLISE FISCAL - ATA.xlsx"}]
    else: 
        return [{"tipo": "ATA", "arquivo": "ANÁLISE FISCAL - ATA.xlsx"},  {"tipo": "CONTRATO", "arquivo": "ANÁLISE FISCAL.xlsx"}] 
    
           
def pegar_numero_protocolo():
    fonte = TXT_PROTOCOLO
    string = ""
    with open(fonte, "r") as txt:
        string = txt.read()
        txt.close()
        return string 
    
def atualizar_numero_protocolo():
    fonte = TXT_PROTOCOLO
    numero_atualizado = int(pegar_numero_protocolo()) + 1
    with open(fonte, "r+") as txt:
        txt.write(str(numero_atualizado))

def pegar_modelos(tipo):
    base_modelos = r"MODELOS BASE\3. MODELOS DE DOCUMENTO"
    if (tipo == "termo"):
        return rf"{ROOT}\{base_modelos}\MODELO DE TERMO.docx"
    elif (tipo == "protocolo"):
        return rf"{ROOT}\{base_modelos}\MODELO DE PROTOCOLO.docx"
    elif (tipo == "portaria"):
        return rf"{ROOT}\{base_modelos}\MODELO DE PORTARIA.docx"
    
def pegar_pasta_downloads():
    return r"C:\Users\Usuario\Downloads"