import xmltodict  # Importa o módulo xmltodict para lidar com arquivos XML
import os  # Importa o módulo os para manipular o sistema de arquivos
import pandas as pd  # Importa o módulo pandas para trabalhar com dataframes
# Importa a classe Workbook do módulo openpyxl
from openpyxl.workbook import Workbook


def pegar_infos(nome_arquivo, valores):
    # Função para extrair informações de um arquivo XML e adicioná-las à lista de valores
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        # Converte o arquivo XML em um dicionário
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if "NFe" in dic_arquivo:
            # Se o dicionário contém a chave "NFe", então a estrutura é diferente
            infos_nf = dic_arquivo["NFe"]['infNFe']
        else:
            # Caso contrário, a estrutura é diferente
            infos_nf = dic_arquivo['nfeProc']["NFe"]['infNFe']
        numero_nota = infos_nf["@Id"]  # Obtém o número da nota fiscal
        # Obtém o nome da empresa emissora
        empresa_emissora = infos_nf['emit']['xNome']
        nome_cliente = infos_nf["dest"]["xNome"]  # Obtém o nome do cliente
        endereco = infos_nf["dest"]["enderDest"]  # Obtém o endereço do cliente
        if "vol" in infos_nf["transp"]:
            # Obtém o peso da mercadoria transportada
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Não informado"  # Se o peso não estiver disponível, define como "Não informado"
        # Adiciona as informações à lista
        valores.append([numero_nota, empresa_emissora,
                       nome_cliente, endereco, peso])


lista_arquivos = os.listdir("nfs")  # Obtém a lista de arquivos na pasta "nfs"

colunas = ["numero_nota", "empresa_emissora", "nome_cliente",
           "endereco", "peso"]  # Define as colunas do dataframe
valores = []  # Lista vazia para armazenar os valores das notas fiscais
for arquivo in lista_arquivos:
    # Chama a função pegar_infos para cada arquivo na lista
    pegar_infos(arquivo, valores)

# Cria um dataframe com as colunas e os valores obtidos
tabela = pd.DataFrame(columns=colunas, data=valores)
# Salva o dataframe em um arquivo Excel chamado "NotasFiscais.xlsx"
tabela.to_excel("NotasFiscais.xlsx", index=False)
