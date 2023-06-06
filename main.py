import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Desativa a verificação do certificado SSL
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# Abrir o arquivo do Excel
workbook = openpyxl.load_workbook('D:/Usuario Barbara/Desktop/conferencia notas fim do mês/FEVEREIRO 2023/teste.xlsx')
print(workbook.sheetnames)
# Selecionar a planilha desejada (substitua 'NomeDaPlanilha' pelo nome correto da planilha)
sheet = workbook['Planilha1']

# URL da página que deseja analisar
url = 'https://www.homehost.com.br/'

# Faz a solicitação HTTP e obtém o conteúdo da página
response = requests.get(url, verify=False)  # verify=False para desabilitar a verificação do certificado

# Cria um objeto BeautifulSoup para analisar o HTML
soup = BeautifulSoup(response.content, 'html.parser')

# Encontra todas as tabelas na página
tables = soup.find_all('table')

# Processa cada tabela conforme necessário
for table in tables:
    # Ler a tabela do navegador e armazenar em um DataFrame (usando a biblioteca pandas)
    data = pd.read_html(str(table))[0]  # Índice [0] para selecionar a primeira tabela encontrada na página

    # Copiar os dados do DataFrame para o Excel
    for row in dataframe_to_rows(data, index=False, header=True):
        sheet.append(row)

# Salvar as alterações no arquivo do Excel
workbook.save('D:/Usuario Barbara/Desktop/conferencia notas fim do mês/FEVEREIRO 2023/REDE.xlsx')

# Fechar o arquivo do Excel
workbook.close()
