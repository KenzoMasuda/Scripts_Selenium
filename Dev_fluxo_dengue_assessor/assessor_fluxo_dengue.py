# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 13:01:40 2023

@author: INFO
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import os
import time


CID_AGRAVOS = ["A90", "A91"]

user = "PHFB"
password = "120499"

diretorio_atual = os.getcwd()
caminho_pasta_download = os.path.join(diretorio_atual, "exportacoes_teste")
if not os.path.isdir(caminho_pasta_download):
    os.mkdir(caminho_pasta_download)
report_download_path = caminho_pasta_download

chromeOptions = webdriver.ChromeOptions()    
prefs = {"download.default_directory" : report_download_path}
chromeOptions.add_experimental_option("prefs", prefs)
chromeOptions.add_argument('--ignore-ssl-errors=yes')
chromeOptions.add_argument('--ignore-certificate-errors')

driver = webdriver.Chrome('chromedriver.exe', chrome_options=chromeOptions)


#----------------------------------------------------------------
#1°Login + MENU - Atendimento - Controle de Agravos
driver.get('https://s56.asp.srv.br/saude.pm.barretos.sp/login')

driver.find_element('id', 'vASPUSUARIOLOGINLOGIN').send_keys(user)
driver.find_element('id', 'vASPUSUARIOLOGINSENHA').send_keys(password)

driver.find_element('id', 'IMGASPLOGIN').click()
time.sleep(2) # try?
driver.find_element('xpath', '//*[@id="ext-gen9"]').click() #click menu
time.sleep(1)
driver.find_element('xpath', '//*[@id="ext-gen93"]').click() #click atendimento
time.sleep(1)
driver.find_element('xpath', '//*[@id="ext-gen148"]').click() #click Controle de Agravos
time.sleep(2)

#driver.switch_to.frame('frame_div_110556') #selecionar frame 'Manutenção Controle de agravos" 


os.startfile("exportacoes_teste", "open")

for agravo in CID_AGRAVOS:
    print("Parou antes de selecionar o frame")
    driver.switch_to.frame('frame_div_110556')
    time.sleep(2)
    print("Parou antes de setar todos")
    Select(driver.find_element(By.XPATH, '//*[@id="vSDCONTROLEAGRAVOSSITUACAO"]')).select_by_visible_text("Todos")
#driver.find_element('xpath', '//*[@id="vCIDAJAX"]').send_keys("A90")
    driver.find_element('id', 'vCIDAJAX').send_keys(agravo)
    time.sleep(2)
    print("Parou antes de selecionar")
    driver.find_element('xpath', '/html/body/div[3]').click()
    time.sleep(2)
    print("Parou antes de procurar")
    driver.find_element('id', 'SEARCHBUTTON').click()
    time.sleep(2)
    print("Parou antes de baixar")
    driver.find_element('xpath', '//*[@id="EXPORT"]').click()
    time.sleep(2)
    print("Parou antes de limpar")
    driver.find_element('id', 'CLEARBUTTON').click()   
    driver.switch_to.default_content()


