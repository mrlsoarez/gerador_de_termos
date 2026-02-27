# Pasta onde ficará armazenado os termos, separados por datas

ROOT = r"C:\Users\Usuario\Documents\MRL"
TXT_PROTOCOLO = r"C:\Program Files (x86)\numero_protocolo\protocolo.txt"

def pegar_endereco_base():
    return rf"{ROOT}\1. ANÁLISE DE PAGAMENTOS\ANÁLISE MENSAL\2026\2. FEVEREIRO"
  
def pegar_planilha_termo(arquivo):
    return rf"{ROOT}\1. ANÁLISE DE PAGAMENTOS\INFO\{arquivo}"
    
def pegar_tipo_termo(prompt):
    if (prompt == "ata"):
        return {
            "corresponde_a": "ata",
            "arquivo": "ANÁLISE FISCAL - ATA.xlsx"
        }
    elif (prompt == "contrato"):
        return {
            "corresponde_a": "contrato",
            "arquivo": "ANÁLISE FISCAL.xlsx"
        }   
    
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
    if (tipo == "termo"):
        return rf"{ROOT}\1. ANÁLISE DE PAGAMENTOS\MODELOS\MODELO DE TERMO.docx"
    elif (tipo == "protocolo"):
        return rf"{ROOT}\1. ANÁLISE DE PAGAMENTOS\MODELOS\MODELO DE PROTOCOLO.docx"