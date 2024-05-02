import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl

# PASSO 0 - CONFIGS
driver = webdriver.Chrome() # SINALIZAR QUE O NAVEGADOR UTILIZAD0 SERÁ O GOOGLE CHROME

# PASSO 1 - ABRIR PLANILHA 
planilha = openpyxl.load_workbook('challenge.xlsx') # AQUI ABRIMOS A PLANILHA QUE CONTÉM AS INFORMAÇÕES PARA CADASTRO

# PASSO 2 - OBTER INFORMAÇÕES 
planilha_work = planilha['Sheet1'] # ATRIBUI A PLANILHA SHEET1 (ONDE ESTÃO AS INFORMAÇÕES) PARA UTILIZARMOS ESSAS INFORMAÇÕESs

# PASSO 3 - ABRIR SITE E INICIAR DESAFIO
driver.get('https://www.rpachallenge.com/') # ABRE O LINK DO DESAFIO
time.sleep(3) # AGUARDA 3 SEGUNDOS PARA A PÁGINA CARREGAR 
driver.find_element(By.XPATH, "//button[contains(text(),'Start')]").click() # CLICA NO BOTÃO START
time.sleep(1) # AGUARDA 1 SEGUNDO PARA INICIAR O DESAFIO

# PASSO 4 - CADASTRAR INFORMAÇÕES
for i in planilha_work.iter_rows(min_row=2, values_only=True):
    if not i[0]:
        continue
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']").send_keys(i[0]) # PREENCHE O FIRST NAME
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']").send_keys(i[1]) # PREENCHE O LAST NAME
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']").send_keys(i[2]) # PREENCHE COMPANY NAME
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']").send_keys(i[3]) # PREENCHE O ROLE COMPANY
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']").send_keys(i[4]) # PREENCHE ADDRESS
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']").send_keys(i[5]) # PREENCHE O EMAIL
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']").send_keys(i[6]) # PREENCHE O TELEFONE
    driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input").click() # CLICA EM SUBMIT

# PASSO 5 - OBTER O RESULTADO E EXIBIR O TEMPO QUE LEVOU PARA CADASTRAR
resultado = driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]").text # OBTEM O RESULTADO
print(resultado)