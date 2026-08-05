# Importação de dados de ambiente, serviços externos e módulos reutilizáveis
from ENV.environment import pegar_endereco_base, pegar_tipo_termo, pegar_planilha_termo, atualizar_numero_protocolo
from SERVICES.PLANILHA import ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS, INICIAR_PLANILHA
from SERVICES.DADOS_EXTERNOS import COLETAR_DADOS_EXTERNOS
from MODULES.GerenciarArquivos import GerenciarArquivos
from MODULES.EncontrarData import EncontrarData

# Importação de Bibliotecas
import os 

def MAIN():
    
    def iniciar_modulo_termos(op):
        
        def realizar_perguntas_iniciais():
            pergunta_inicial = CAPTURAR_RESPOSTA(
                "Bem vindo ao gerador de termos! Escolha dentre as opções para gerar: \n1. Contrato\n2. Ata\n-> ", ("1", "2")
            )
            pergunta_update = CAPTURAR_RESPOSTA(
                "Deseja atualizar os dados da planilha? (S/N) -> ", ("s", "n")
            )
            pergunta_total = CAPTURAR_RESPOSTA(
                "Deseja uma geração total? (S/N) -> ", ("s", "n")
            )
            return pergunta_inicial, pergunta_update, pergunta_total
                
        def criar_pasta_termos(gerenciador):
            
            mes_numero = EncontrarData("mes", False)
            mes_extenso = EncontrarData("mes", True) 
            dia = EncontrarData("dia")
        
            os.chdir(gerenciador.pasta_base)

            PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
            PASTA_DIA = f"{dia}-{mes_numero}"
            PROTOCOLO = f"REMESSA X - PROTOCOLO N° {gerenciador.numero_protocolo}"
        
            gerenciador.criarPasta(PASTA_MES, True)
            gerenciador.criarPasta(PASTA_DIA, True)
            gerenciador.criarPasta(PROTOCOLO, True)
            gerenciador.criarPasta("PROTOCOLOS", False)
            gerenciador.criarPasta("TERMOS", True)
            gerenciador.criarPasta(gerenciador.tipo_arquivo, True)
            gerenciador.criarPasta("WORD")
            gerenciador.criarPasta("PDF")
        
            gerenciador.pasta_atual = rf"{gerenciador.pasta_base}\{PASTA_MES}\{PASTA_DIA}\{PROTOCOLO}\TERMOS\{gerenciador.tipo_arquivo}"
           
        def capturar_dados_externos(option):
            print("Iniciando a coleta de dados externos.. por favor, aguarde..")
            try:
                if (option == "1"): dados_contratos = COLETAR_DADOS_EXTERNOS("contratos")
                #dados_servidores =  COLETAR_DADOS_EXTERNOS("servidores")
                #dados_empenhos = COLETAR_DADOS_EXTERNOS("empenhos")
                #dados_liquidacao =  COLETAR_DADOS_EXTERNOS("liquidacoes")
                pass
            except: 
                print("Algo deu errado no processo de busca de dados")
            else:
                print("Dados encontrados e armazenados para inserção nas planilhas")

            print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
            print("Inserindo as informações na planilha de análise base.. favor, aguardar.")     
            try: 
                if (option == "1"): ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Base", dados_contratos)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Servidores", dados_servidores)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Empenhos", dados_empenhos)
                #ATUALIZAR_PLANILHA_COM_DADOS_EXTERNOS(PASTA_PLANILHA_ANALISE, "Liquidacoes", dados_liquidacao)
            except:
                print("Algo deu errado no processo de inserir as informações nas planilhas")
            else: 
                print("Dados inseridos nas planilhas.")
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
        
        resposta = realizar_perguntas_iniciais()
            
        pergunta_inicial = resposta[0][0]
        pergunta_update = resposta[1][0]
        pergunta_total = resposta[2][0]
            
        #pergunta_inicial = "2"
        #pergunta_update = "s"
        #pergunta_total = "n"
                    
        TIPO_TERMO = pegar_tipo_termo(pergunta_inicial)[0]
        PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO["arquivo"])
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), TIPO_TERMO["tipo"])
            
        criar_pasta_termos(GERENCIADOR_PASTAS)
        
        print(GERENCIADOR_PASTAS.numero_protocolo)
        if (pergunta_update == "s"): 
            capturar_dados_externos(pergunta_inicial)

        parametros = {
            "op": op, 
            "modo_total": True if pergunta_total == "s" else False
        }
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■\nIniciando planilha..")
        INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)

    def iniciar_modulo_protocolo(op):
        
        print("■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■")
        
        def entrar_na_pasta_de_protocolo():
            mes_numero = EncontrarData("mes", False)
            mes_extenso = EncontrarData("mes", True) 
            dia = EncontrarData("dia")       
            
            PASTA_MES = f"{mes_numero[1:]}. {mes_extenso}"
            PASTA_DIA = f"{dia}-{mes_numero}"
            
            return  GERENCIADOR_PASTAS.entrarEmPasta(fr"{PASTA_MES}\{PASTA_DIA}\REMESSA X - PROTOCOLO N° {GERENCIADOR_PASTAS.numero_protocolo}\PROTOCOLOS")
            
        def gerar_protocolo(pergunta_inicial):
 
            ARRAY = []
            TIPO_TERMO = pegar_tipo_termo(pergunta_inicial)
            
            parametros = {
                "op": op, 
                "ARR": ARRAY
            }
            
            for i in range(len(TIPO_TERMO)):
                PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO[i]["arquivo"])
                GERENCIADOR_PASTAS.set_tipo_arquivo = TIPO_TERMO[i]["tipo"]
                ARRAY = INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)
            
            print(ARRAY)
            protocolo = ARRAY[0]
            
            protocolo.copiar_arquivo(protocolo.modelo, protocolo.endereco)
            protocolo.criar_arquivo(ARRAY)
            
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), "")
        
        if (not entrar_na_pasta_de_protocolo()):
            print("Pasta de relatório ainda não existe!")
            return 
        
        pergunta_inicial = CAPTURAR_RESPOSTA(
            "Deseja gerar relatórios de quais informações?: \n1. Contrato\n2. Ata\n3. Ambos\n-> ", ("1", "2", "3")
        )
        
        gerar_protocolo(pergunta_inicial)
        
        
      
        
    
         
    
        """
        
        if (pergunta_inicial[0] == "3"):
            for i in range(1, 3):
                TIPO_TERMO = pegar_tipo_termo(str(i))
        
        
        TIPO_TERMO = pegar_tipo_termo(pergunta_inicial)
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(), pergunta_inicial)
        """
        pass 
    
    def iniciar_modulo_portaria(op):
        parametros = {
            "op": op
        }
        TIPO_TERMO = pegar_tipo_termo("2")
        PASTA_PLANILHA_ANALISE = pegar_planilha_termo(TIPO_TERMO[0]["arquivo"])
        GERENCIADOR_PASTAS = GerenciarArquivos(pegar_endereco_base(op), "")
        INICIAR_PLANILHA(PASTA_PLANILHA_ANALISE, GERENCIADOR_PASTAS, parametros)
        pass 
    
    while True:        
        resposta = CAPTURAR_RESPOSTA("Bem vindo! Escolha dentre as opções \n1. Gerar termos aditivos\n2. Gerar relatório\n3. Gerar portaria\n4. Atualizar número de protocolo\n5. Encerrar\n-> ", ("1", "2", "3", "4", "5"))
        opcao = resposta[0]
        if (opcao == "1"):
            iniciar_modulo_termos(opcao)
        elif (opcao == "2"):
            iniciar_modulo_protocolo(opcao)
        elif (opcao == "3"):
            iniciar_modulo_portaria(opcao)
        elif (opcao == "4"):
            atualizar_numero_protocolo()
        elif (opcao == "5"):
            print("Encerrando.")
            break
    
def CAPTURAR_RESPOSTA(mensagem, dado_esperado):

    PERGUNTA = input(mensagem).lower()

    while PERGUNTA not in dado_esperado:
        print(("Opção não válida, por favor, digite uma das opções ao lado ", dado_esperado))
        PERGUNTA = input(mensagem).lower()
        
    return (PERGUNTA, True)

MAIN()