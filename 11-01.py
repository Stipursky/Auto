from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
import time

s = Service(r'C:\Users\Matheus\Desktop\Hash PY\cgeckodriver.exe')
navegador = webdriver.Firefox(service=s)

navegador.get('https://www.google.com/')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Cotação dólar')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(4)
c_dolar = navegador.find_element('xpath', '/html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/span[1]').get_attribute('data-value')

print(c_dolar)

navegador.get('https://www.google.com/')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Cotação euro')
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
time.sleep(4)
c_euro = navegador.find_element('xpath', '/html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/span[1]').get_attribute('data-value')

print(c_euro)

navegador.get('https://www.melhorcambio.com/ouro-hoje')
c_ouro = navegador.find_element('xpath','//*[@id="comercial"]').get_attribute('value')
c_ouro = c_ouro.replace(',','.')

print(c_ouro)

navegador.quit()

import pandas as pd

tabela = pd.read_excel(r'C:\Users\Matheus\Desktop\Hash PY\Aula 3\a\Produtos.xlsx')

tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(c_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(c_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(c_ouro)

tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']
print(tabela)

tabela.to_excel("Produtos - Novos.xlsx", index=False)
