#Passo 1: Importando as Bibliotecas
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import pandas as pd
import openpyxl

# Passo 2: Abrindo o Navegador
navegador = webdriver.Chrome(executable_path='chromedriver.exe')

navegador.get('https://www.google.com/')

# Passo 3: Pesquisando a cotação do Dólar
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Cotação Dólar') # Escrevendo o nome
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER) #clicando no enter
cotacao_Dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value') #Selecionando um atributo 
print(cotacao_Dolar)

#Passo 4: Pesquisando a cotação do Euro

navegador.get('https://www.google.com/') 
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')# Escrevendo o nome
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER) #Clicando no Enter
cotacao_Euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value') #Selecionando um Atributo
print(cotacao_Euro)

#Passo 5: Pesquisando a cotação do Ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')
cotacao_Ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacao_Ouro = cotacao_Ouro.replace(',','.')
print(cotacao_Ouro)

#Passo 6: Atualizar a minha base de dados com as novas cotações

tabela = pd.read_excel('Produtos.xlsx')

#Atualizar a cotação de acordo com a moeda correspondente

'''
 -->Tabela.loc[Linha,Coluna]
 -->As linhas onde a coluna 'moeda' é igual a moeda correspondente será alterada
'''

#Dolar: Na coluna moeda, se linha for igual'Dólar', a coluna 'Cotação' recebe o float da variável cotacao_Dolar.
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_Dolar)

#Euro: Na coluna moeda, se linha for igual 'Euro', a coluna 'cotação' recebe o float da variável cotacao_Euro.
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_Euro)

#Ouro: Na coluna moeda, se linha for igual 'Ouro', a coluna 'cotação' recebe o float da variável cotacao_Ouro.
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_Ouro)

#Atualizar o preço de compra = Preço Original * Cotação
tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']

#Atualizar o preço de venda = preço de compra * margem
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

print(tabela)

tabela.to_excel('Produtos Novos.xlsx', index = False) #O parâmetro False serve para tirar o indice para quando exportar essa tabelas.
navegador.quit()