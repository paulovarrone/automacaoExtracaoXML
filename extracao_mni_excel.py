from zeep import Client
from zeep.transports import Transport
import requests
from datetime import datetime
import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

excel_file = './proc.xlsx'
df = pd.read_excel(excel_file)
processos_xcel = df['processos'].tolist()

wsdl = os.getenv('WSDL')

transport = Transport(session=requests.Session())

client = Client(wsdl=wsdl, transport=transport)

id_consultante = os.getenv('ID_CONSULTANTE')
senha_consultante = os.getenv('SENHA_CONSULTANTE')

def consultar_processo(numero_processo):
    request_data = {
        'idConsultante': id_consultante,
        'senhaConsultante': senha_consultante,
        'numeroProcesso': numero_processo,
        'movimentos': True,
        'incluirCabecalho': True,
        'incluirDocumentos': False
    }

    response = client.service.consultarProcesso(**request_data)
    return response

def extrair_info_partes(polo):
    partes_info = []
    for parte in polo.parte:
        pessoa = parte.pessoa
        nome = pessoa.nome
        documento_principal = pessoa.numeroDocumentoPrincipal
        advogados = list(map(lambda advogado: advogado.nome, parte.advogado))
        partes_info.append({
            'nome': nome,
            'documento_principal': documento_principal,
            'advogados': advogados
        })
    return partes_info

def pegar_polo(dados_basicos, polo_tipo):
    for polo in dados_basicos.polo:
        if polo.polo == polo_tipo:
            return extrair_info_partes(polo)
    return []

# Lista para armazenar dados dos processos
dados_processos = []
cont = 0
for numero_processo in processos_xcel:
    try:
        
        print(f'Processando número do processo: {numero_processo} {cont}')
        response = consultar_processo(numero_processo)
        dados_basicos = response.processo.dadosBasicos
        nome_orgao = dados_basicos.orgaoJulgador.nomeOrgao
        valor_causa = dados_basicos.valorCausa

        data_autuacao = dados_basicos.dataAjuizamento
        data_obj = datetime.strptime(data_autuacao, "%Y%m%d%H%M%S")
        data_autuacao = data_obj.strftime("%d/%m/%Y")

        ultimo_movimento = response.processo.movimento[0]
        movimento_nacional = ultimo_movimento.movimentoNacional
        complemento = " ".join(movimento_nacional.complemento) 

        partes_at = pegar_polo(dados_basicos, 'AT')
        partes_pa = pegar_polo(dados_basicos, 'PA')

        # Inicializando as variáveis para concatenar as informações dos autores e réus
        autor_nomes = []
        autor_documentos = []
        autor_advogados = []

        reu_nomes = []
        reu_documentos = []
        reu_advogados = []

        for parte in partes_at:
            autor_nomes.append(parte['nome'])
            autor_documentos.append(parte['documento_principal'])
            autor_advogados.extend(parte['advogados'])

        for parte in partes_pa:
            reu_nomes.append(parte['nome'])
            reu_documentos.append(parte['documento_principal'])
            reu_advogados.extend(parte['advogados'])

        dados_processos.append({
            'Numero do Processo': numero_processo,
            'Nome do Órgão': nome_orgao,
            'Valor da Causa': valor_causa,
            'Data de Autuação': data_autuacao,
            'Última Movimentação': complemento,
            'Autor Nome': ', '.join(autor_nomes),
            'Autor Documento Principal': ', '.join(autor_documentos),
            'Autor Advogados': ', '.join(autor_advogados),
            'Réu Nome': ', '.join(reu_nomes),
            'Réu Documento Principal': ', '.join(reu_documentos),
            'Réu Advogados': ', '.join(reu_advogados)
        })

        cont += 1

    except Exception as e:
        print(f'Erro ao processar o número do processo {numero_processo}: {e}')

# Convertendo lista de dicionários para DataFrame
df_processos = pd.DataFrame(dados_processos)

# Diretório onde o arquivo Excel será salvo
diretorio = './Relatorio'

# Certificando-se de que o diretório existe
if not os.path.exists(diretorio):
    os.makedirs(diretorio)

# Caminho completo do arquivo Excel
caminho_arquivo = os.path.join(diretorio, 'Processos.xlsx')

# Salvando DataFrame em um arquivo Excel
df_processos.to_excel(caminho_arquivo, index=False)
print(f'Dados salvos em {caminho_arquivo}')
