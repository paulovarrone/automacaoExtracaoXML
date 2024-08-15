from zeep import Client
from zeep.transports import Transport
import requests
from datetime import datetime
from fpdf import FPDF
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
    

def criar_pdf(numero_processo, nome_orgao, valor_causa, data_autuacao, complemento, parte_at, parte_pa):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    def add_texto(text):
        pdf.multi_cell(0, 10, text)

    add_texto(f'Numero do Proesso: {numero_processo}')
    add_texto(f"Nome do Órgão: {nome_orgao}")
    add_texto(f"Valor da Causa: {valor_causa}")
    add_texto(f'Data Autuação: {data_autuacao}')
    add_texto(f'Ultima Movimentação: {complemento}')

    add_texto("\nPartes AUTOR:")
    for parte in parte_at:
        add_texto(f"Nome: {parte['nome']}")
        add_texto(f"CPF: {parte['documento_principal']}")
        add_texto(f"Advogados: {', '.join(parte['advogados'])}")
        add_texto("------")

    add_texto("\nPartes RÉU:")
    for parte in parte_pa:
        add_texto(f"Nome: {parte['nome']}")
        add_texto(f"CNPJ: {parte['documento_principal']}")
        add_texto(f"Advogados: {', '.join(parte['advogados'])}")
        add_texto("------")

    return pdf

def salvar_pdf(pdf, numero_processo):
    nome_arquivo = f'{numero_processo}.pdf'
    diretorio = './Processos'
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    if not os.path.exists(diretorio):
        os.makedirs(diretorio)

    if os.path.exists(caminho_completo):
        print(f'ARQUIVO JÁ EXISTE: {caminho_completo}')
    else:
        pdf.output(caminho_completo)
        print(f'PDF salvo com sucesso: {caminho_completo}')


for numero_processo in processos_xcel:
    try:
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

        parte_at = pegar_polo(dados_basicos, 'AT')
        parte_pa = pegar_polo(dados_basicos, 'PA')

        pdf = criar_pdf(numero_processo, nome_orgao, valor_causa, data_autuacao, complemento, parte_at, parte_pa)
        salvar_pdf(pdf, numero_processo)
    except Exception as e:
        print(f'Erro ao processar o número do processo {numero_processo}: {e}')