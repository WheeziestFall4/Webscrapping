
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from bs4 import BeautifulSoup
from urllib.request import urlopen
import pandas as pd
import openpyxl
import os.path
import wmi

navegador = webdriver.Chrome()

# Dicionario de sites para buscar
sites = {
    'Amazon': 'https://www.amazon.com.br/PlayStation-CFI-1214A01X-Console-5/dp/B0BNSR3MW9/ref=asc_df_B0BNSR3MW9/?tag=googleshopp00-20&linkCode=df0&hvadid=431423537910&hvpos=&hvnetw=g&hvrand=11025709941912363269&hvpone=&hvptwo=&hvqmt=&hvdev=c&hvdvcmdl=&hvlocint=&hvlocphy=1001729&hvtargid=pla-1943883661613&th=1',
    'Submarino': 'https://www.submarino.com.br/produto/5315680611?pfm_carac=ps5&pfm_index=2&pfm_page=search&pfm_pos=grid&pfm_type=search_page&offerId=62daf4da6e9e5a65df43d6a6&cor=BRANCO&voltagem=BIVOLT&condition=NEW'
}

# Guarda o class do nome e do preco e o tipo colocado no html
# [tipoNome, nome, tipoPreco, preco]
buscar = {
    'Amazon': [".a-price-whole",".a-size-large.product-title-word-break"],
    'Submarino':[".src__BestPrice-sc-1jnodg3-5.ykHPU.priceSales",".src__Title-sc-1xq3hsd-0.eEEsym"]
    # 'Amazon': ["span", "a-size-large product-title-word-break", "span", "a-price-whole"],
    # 'Submarino': ["h1", "src__Title-sc-1xq3hsd-0 eEEsym", "span", "src__BestPrice-sc-1jnodg3-5 ykHPU priceSales"]
}

precos = {
    'Amazon': '',
    'Submarino': ''
}

nomes = {
    'Amazon': '',
    'Submarino': ''
}

# Verificar se existe arquivo TabelaPrecos
if not os.path.isfile('TabelaPrecos.xlsx'):

    workbook = xlsxwriter.Workbook('TabelaPrecos.xlsx')
    print("Criando arquivo excel Tabela de preços")
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Loja')
    worksheet.write('B1', 'Nome')
    worksheet.write('C1', 'Preco')
    workbook.close()
    # Editar o arquivo Excel
    # variavel_edicao = openpyxl.load_workbook('TabelaPrecos.xlsx')
    # abaSheet1 = variavel_edicao['Sheet1']

x = 0
for chave,valor in sites.items():
    print(valor)
    navegador.get(valor)
    time.sleep(3)
    print(navegador.find_element(By.CSS_SELECTOR, buscar[chave][0]).text)
    print(navegador.find_element(By.CSS_SELECTOR, buscar[chave][1]).text)
    precos[chave] = str(navegador.find_element(By.CSS_SELECTOR, buscar[chave][0]).text).strip()
    nomes[chave] = str(navegador.find_element(By.CSS_SELECTOR, buscar[chave][1]).text).strip()
    #html = urlopen(valor)
    # soup = BeautifulSoup(html, 'html.parser')
    # precoUnico = soup.find(buscar[chave][2], {'class': buscar[chave][3]})
    # nomeProdutoUnico = soup.find(buscar[chave][0], {'class': buscar[chave][1]})

# O arquivo existindo deve verificar se ele esta aberto
f = wmi.WMI()

for process in f.Win32_Process():
    if "EXCEL.EXE" == process.Name:
        print("Aplication is running")
        os.system('taskkill /f  /im EXCEL.EXE')
        # os.close(os.path('TabelaPrecos.xlsx'))
        print("Closing")
        print("Aplication is not running anymore")
        break


df_final = pd.read_excel('TabelaPrecos.xlsx')
print(df_final)

for chave in sites.keys():
    # Cria nova tabela organizacional em colunas
    nova_df = pd.DataFrame({"Loja": [chave], "Nome": [nomes[chave]], "Preco": [precos[chave]]})
    print(nova_df)
    df_final = pd.concat([df_final, nova_df], ignore_index=True)
    # df_final = pd.concat([df_lida, nova_linha])

# Apagar arquivo excel antigo
os.remove('TabelaPrecos.xlsx')

# Converte para um arquivo Excel
df_final.to_excel('TabelaPrecos.xlsx', index=False)
print("Criando arquivo excel Tabela de preços")