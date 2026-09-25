import os 
from openpyxl import load_workbook
from pathlib import Path
from datetime import date, timedelta
import shutil

from Services.Calendario import EncontrarData 

# Variável de preparação do ambiente inicial, criando pastas

##########################

# Ao instanciar um protocolo, instanciar imediatamente o protocolo anterior. 
    # QUAIS SÃO AS INFORMAÇÕES PRINCIPAIS DO PROTOCOLO ANTERIOR?
        # Númeração e pasta de armazenamento 
# Ou seja, ao settar o número de protocolo, também preciso setar o protocolo anterior, encontrando
 # sua respectiva númeração e data, que não necessariamente vai ser sequencial
class Env: 
    
    def __init__(self):
        self.pastaBase = rf"{Path.home()}\Documents\MRL2"
        self.pastaPlanilha = rf"{self.pastaBase}\MODELOS BASE\1. PLANILHAS - ANÁLISE FISCAL - ATA & CONTRATOS"
        self.pastaProtocolo = rf"{self.pastaBase}\MODELOS BASE\2. CONTROLE PROTOCOLO"
        self.pastaModelo = rf"{self.pastaBase}\MODELOS BASE\3. MODELOS DE DOCUMENTOS"
        self.pastaAtual = self.pastaBase 
        self.protocoloAntigo = None
        pass 

    def criarPasta(self, pasta, navegar=False):
        operacoesOs = OperacoesOs(self)
        if operacoesOs.criarPasta(pasta):
            if navegar:
                self.pastaAtual = pasta

    def criarPastasIniciais(self):
        self.criarPasta(self.pastaBase)
        self.criarPasta(self.pastaPlanilha)
        self.criarPasta(self.pastaProtocolo)
        self.criarPasta(self.pastaModelo)

    def criarPastasDatas(self): 
        self.criarPasta(rf"{self.pastaAtual}\1. ANÁLISE DE PAGAMENTOS", navegar=True)
        self.criarPasta(rf"{self.pastaAtual}\{EncontrarData('ano')}", navegar=True)
        self.criarPasta(rf"{self.pastaAtual}\{EncontrarData('mes', False)}. {EncontrarData('mes', True)}", navegar=True)
        self.criarPasta(rf"{self.pastaAtual}\{EncontrarData('dia')}-{EncontrarData('mes', False)}", navegar=True)
        self.pastaHoje = self.pastaAtual 
        
    def setGerenciador(self):
        self.gerenciador = OperacoesOs(self)
    
    def setProtocolo(self):
        
        try: 
            txtProtocolo = rf"{self.pastaProtocolo}\protocolo.txt"
            with open(txtProtocolo, "r") as txt:
                string = txt.read()
                txt.close()
                self.numeroProtocolo = string 
        except Exception as e:
            print(e)
        else: 
            print("-> Protocolo encontrado!")
                
    def setProtocoloAntigo(self): 
        
        def encontrarProtocoloAntigo(pasta):
            count = 1
            while count < 11: 
                caminhoProtocolo = rf"{pasta}\REMESSA - PROTOCOLO {int(self.numeroProtocolo)-count}"
                if (self.gerenciador.pastaExiste(caminhoProtocolo)): return caminhoProtocolo
                count += 1
                
        pastasProtocolosHoje = self.gerenciador.listarArquivos(self.pastaHoje)

        if (len(pastasProtocolosHoje) > 1): 
            self.protocoloAntigo = encontrarProtocoloAntigo(self.pastaHoje)
    
    def setPastaAnterior(self): 
        pastasDatas = self.gerenciador.listarArquivos(self.pastaHoje.rsplit("\\", 1)[0])
        diaHoje = date.today().day 
        pastaAnterior = ""
        for i in range(len(pastasDatas), 0, -1): 
            if pastasDatas[i-1][:2] < str(diaHoje): 
                pastaAnterior = pastasDatas[i-1]
                break
        self.pastaAnterior = rf"{self.pastaHoje.rsplit("\\", 1)[0]}\{pastaAnterior}"
        pass 
    
    # Fazer a distinção entre protocolo anterior e novo, copiando os mantidos na planilha para
        # a nova pasta e com a ordem rearranjada (organizar tudo na pasta anterior também)
    def copiarTermosAnteriores(self, dados): 

        protocolo = f"REMESSA - PROTOCOLO {self.numeroProtocolo}"
        for chave in dados:
            pastaNova = rf"{self.pastaAtual}\{chave.upper()}"
            pastaAntiga = rf"{self.protocoloAntigo}\TERMOS\{chave.upper()}"
            arquivos = self.gerenciador.listarArquivos(pastaAntiga)
            for index in range(len(dados[chave])): 
                dado = dados[chave][index] 
                nomeArquivo = f"{dado.contratado} - AF {dado.af[:4]}"
                for arq in arquivos: 
                    if (nomeArquivo in arq): 
                        ext = os.path.splitext(arq)[1]
                        copiarArquivoAntigo = rf"{pastaAntiga}\{arq}"
                        copiarArquivoNovo = rf"{pastaNova}\{index+1}. {nomeArquivo}{ext}"                   
                        if self.gerenciador.arquivoExiste(rf"{copiarArquivoAntigo}"): 
                            self.gerenciador.copiarArquivo(copiarArquivoAntigo, copiarArquivoNovo)
                            self.gerenciador.removerArquivo(copiarArquivoAntigo)
            self.gerenciador.corrigirOrdemArquivos(pastaAntiga)
        """
        def encontrarDiaAnterior(): 
            data = date.today() - timedelta(days=1)
            while data.weekday() >= 5:
                data -= timedelta(days=1)
            return data 
        
        def encontrarProtocoloAnterior():
            operacoesOs = OperacoesOs(self)
            count = 1
            pasta = self.pastaHoje 
            while True: 
                caminhoProtocolo = rf"{pasta}\REMESSA - PROTOCOLO {int(self.numeroProtocolo)-count}"
                if (operacoesOs.pastaExiste(caminhoProtocolo)): return caminhoProtocolo
                count += 1 
                if (count == 10):
                    pasta = self.pastaOntem
                    count = 1
                    
        dataOntem = encontrarDiaAnterior()
        
        mes = f"{dataOntem.month:02d}"
        dia = f"{dataOntem.day:02d}"

        if arq != False: 
            pastaMes = self.pastaAtual.rsplit("\\", 3)[0]
            self.pastaOntem = rf"{pastaMes}\{dia}-{mes}"
            self.pastaOntem = encontrarProtocoloAnterior()
        else:
            self.pastaOntem = rf"{self.pastaAtual.rsplit('\\', 1)[0]}\{dia}-{mes}\{protocolo}"

        if operacoesOs.existe(self.pastaOntem):
            if arq != False: 
                for chave in arq: 
                    for i in range(len(arq[chave])): 
                        novoDiretorio = rf"{self.pastaAtual}\{chave.upper()}"
                        diretorioAntigo = rf"{self.pastaOntem}\TERMOS\{chave.upper()}"
                        nomeArquivo = f"{arq[chave][i].contratado} - AF {arq[chave][i].af[:4]}"
                        self.verificarArquivo(diretorioAntigo, novoDiretorio, nomeArquivo, i)
            else: 
                try:
                    operacoesOs.copiarPasta(self.pastaOntem, rf"{self.pastaAtual}\{protocolo}")
                except Exception as e:
                    print(e) 
                finally: 
                    self.pastaAtual = rf"{self.pastaAtual}\{protocolo}\TERMOS"
                    return True

        """
    
    # Copiar protocolo anterior, a partir da última data anterior encontrada. O protocolo permanece o atual,
        # apenas transportado para o dia atual
    def copiarProtocoloAnterior(self): 
        protocolo = f"REMESSA - PROTOCOLO {self.numeroProtocolo}"
        pastaAnterior = rf"{self.pastaAnterior}\{protocolo}"
        if self.gerenciador.pastaExiste(pastaAnterior):
            try:
                self.gerenciador.copiarPasta(pastaAnterior, rf"{self.pastaAtual}\{protocolo}")
            except Exception as e:
                print(e) 
            finally: 
                self.pastaAtual = rf"{self.pastaAtual}\{protocolo}\TERMOS"
                self.gerenciador.removerPasta(pastaAnterior)
                return True
    
    def criarPastasProtocolos(self):
        
        verificacao = self.copiarProtocoloAnterior()          
        if verificacao: 
            return 
        
        self.criarPasta(rf"{self.pastaAtual}\REMESSA - PROTOCOLO {self.numeroProtocolo}", navegar=True)
        self.criarPasta(rf"{self.pastaAtual}\PROTOCOLOS")
        self.criarPasta(rf"{self.pastaAtual}\TERMOS", navegar=True)
        self.criarPasta(rf"{self.pastaAtual}\ATA")
        self.criarPasta(rf"{self.pastaAtual}\CONTRATO")

    def setPlanilha(self):
        try: 
            self.caminhoPlanilha = rf"{self.pastaPlanilha}\ANÁLISE FISCAL - Copia.xlsx"
            planilha = load_workbook(rf"{self.caminhoPlanilha}", data_only=True)
        except Exception as e: 
            print(e)
        else: 
            self.planilha = planilha
            print(f"-> Planilha de análise fiscal encontrada: {self.planilha}")

    def incrementarProtocolo(self):  
        fonte = rf"{self.pastaProtocolo}\protocolo.txt"
        numeroAtualizado = int(self.numeroProtocolo) + 1
        with open(fonte, "r+") as txt:
            txt.write(str(numeroAtualizado))
    
    def setModelo(self, op):
        modelos = {
            1: self.pastaModelo + r"\MODELO DE TERMO.docx",
            2: self.pastaModelo + r"\MODELO DE PROTOCOLO.docx"
        }
        self.modelo = modelos[op]

    def verificarArquivo(self, caminho, caminhoNovo, nome, ordem):
        
        #arquivos = operacoesOs.listarArquivos(caminho) 
        """
        
        for arq in arquivos: 
            if nome in arq: 
                ext = os.path.splitext(arq)[1]
                pastaOrigem = rf"{caminho}\{arq}"
                pastaNova = rf"{caminhoNovo}\{ordem+1}. {nome}{ext}"
                print(pastaOrigem, pastaNova)
                if operacoesOs.copiarArquivo(pastaOrigem, pastaNova):
                    operacoesOs.removerArquivo(pastaOrigem)
        
        operacoesOs.corrigirOrdemArquivos(caminho)

        """
class OperacoesOs:

    def __init__(self, env):
        self.env = env

    def existe(self, caminho):
        try:
            return os.path.exists(caminho)
        except Exception as e:
            print(e)
            return False

    def arquivoExiste(self, caminho):
        try:
            return os.path.isfile(caminho)
        except Exception as e:
            print(e)
            return False

    def pastaExiste(self, caminho):
        try:
            return os.path.isdir(caminho)
        except Exception as e:
            print(e)
            return False

    def criarPasta(self, caminho):
        try:
            os.makedirs(caminho, exist_ok=True)
            return True
        except Exception as e:
            print(e)
            return False

    def listarArquivos(self, caminho):
        try:
            return os.listdir(caminho)
        except Exception as e:
            print(e)
            return []
    
    def corrigirOrdemArquivos(self, caminho):
        try:
            arquivos = self.listarArquivos(caminho)
            grupos = {}
            for arquivo in arquivos:
                nome, extensao = os.path.splitext(arquivo)
                if ". " in nome:
                    nome = nome.split(". ", 1)[1]
                if nome not in grupos:
                    grupos[nome] = []
                grupos[nome].append((arquivo, extensao))
            temporarios = []
            for i, (nome, arquivosGrupo) in enumerate(grupos.items(), 1):
                for j, (arquivo, extensao) in enumerate(arquivosGrupo):
                    caminhoAntigo = rf"{caminho}\{arquivo}"
                    temporario = rf"{caminho}\__temp_{i}_{j}{extensao}"
                    os.rename(caminhoAntigo, temporario)
                    temporarios.append((temporario, rf"{caminho}\{i}. {nome}{extensao}"))
            for temporario, nomeNovo in temporarios:
                os.rename(temporario, nomeNovo)
        except Exception as e:
            print(e)
            
    def copiarArquivo(self, origem, destino):
        try:
            shutil.copy(origem, destino)
            return True
        except Exception as e:
            print(e)
            return False

    def copiarPasta(self, origem, destino):
        try:
            shutil.copytree(origem, destino)
            return True
        except Exception as e:
            print(e)
            return False
    
    def removerPasta(self, origem):
        try:
            shutil.rmtree(origem)
            return True
        except Exception as e:
            print(e)
            return False
            
    def removerArquivo(self, caminho):
        try:
            os.remove(caminho)
            return True
        except Exception as e:
            print(e)
            return False
        

def inicializarAmbiente():
    
    controladorAmbiente = Env()
    controladorAmbiente.setGerenciador()
    
    controladorAmbiente.setProtocolo()
    
    controladorAmbiente.criarPastasIniciais()
    controladorAmbiente.criarPastasDatas()
    
    controladorAmbiente.setPastaAnterior()
    controladorAmbiente.setProtocoloAntigo()
    
    controladorAmbiente.criarPastasProtocolos()
    
    controladorAmbiente.setPlanilha()
    
    return controladorAmbiente
