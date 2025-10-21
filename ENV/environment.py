# Pasta onde ficará armazenado os termos, separados por datas
def pegar_endereco_base():
    return r"C:\Users\Usuario\Documents\MURILO\3. ANÁLISES DE NOTAS FISCAIS\1. ANÁLISE MENSAL\10. OUTUBRO"
  


def pegar_planilha_termo(arquivo):
    return rf"C:\Users\Usuario\Documents\MURILO\3. ANÁLISES DE NOTAS FISCAIS\2. GERADOR\INFO\{arquivo}"
    
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
    source = r"C:\Program Files (x86)\numero_protocolo\protocolo.txt"
    string = ""
    with open(source, "r") as txt:
        string = txt.read()
        txt.close()
        return string 
