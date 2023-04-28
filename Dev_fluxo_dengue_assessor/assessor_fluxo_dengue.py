# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 13:01:40 2023

@author: KENZO
"""
def _agravos_dengue_assessor():
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import Select
    import os
    import time
    from glob import glob
    import pandas as pd
    from pathlib import Path
    
    def remove_old_files(path, extension=None):
        for file in os.listdir(path):
            if extension:
                if Path(path, file).suffix == extension:
                    os.remove(Path(path, file))
            else:
                os.remove(Path(path, file))
                
    def download_agravos(agravos):    
        for agravo in agravos:
            again = True
            while again:
                try:
                    driver.switch_to.frame('frame_div_110556')
                    time.sleep(2)
                    Select(driver.find_element(By.XPATH, '//*[@id="vSDCONTROLEAGRAVOSSITUACAO"]')).select_by_visible_text("Todos")
                #driver.find_element('xpath', '//*[@id="vCIDAJAX"]').send_keys("A90")
                    driver.find_element('id', 'vCIDAJAX').click()
                    driver.find_element('id', 'vCIDAJAX').send_keys(agravo)
                    time.sleep(2)
                    #driver.find_element('xpath', '/html/body/div[3]').click()
                    driver.find_element('xpath', '//*[@id="ui-active-menuitem"]/div/div[2]').click()
                    time.sleep(2)
                    driver.find_element('id', 'SEARCHBUTTON').click()
                    time.sleep(2)
                    driver.find_element('xpath', '//*[@id="EXPORT"]').click()
                    time.sleep(2)
                    driver.find_element('id', 'CLEARBUTTON').click()   
                    driver.switch_to.default_content()
                    again = False
                except:
                    print('aqui')
                    again = True
            
    def concat_agravos_dengue_assessor(path):
        list_agravos = glob(path + "\*.xls")
        
        df = []
    
        for agravos in list_agravos:
            df.append(pd.read_excel(agravos))
            
        df_dengue_assessor = pd.concat([df[0], df[1]])
        df_dengue_assessor.to_excel(path + "\\" + "DENGUE2.xlsx", index=False)
    
    CID_AGRAVOS = ["A90", "A91"]
    
    user = os.environ.get('login_assessor')
    password = os.environ.get('pass_assessor')
    
    diretorio_atual = os.getcwd()
    default_download = os.path.join(diretorio_atual, "exportacoes_teste")
    if not os.path.isdir(default_download):
        os.mkdir(default_download)
    
    chromeOptions = webdriver.ChromeOptions()    
    prefs = {"download.default_directory" : default_download}
    chromeOptions.add_experimental_option("prefs", prefs)
    chromeOptions.add_argument('--ignore-ssl-errors=yes')
    chromeOptions.add_argument('--ignore-certificate-errors')
    
    driver = webdriver.Chrome('chromedriver.exe', chrome_options=chromeOptions)
    
    #----------------------------------------------------------------
    #1°Login + MENU - Atendimento - Controle de Agravos
    driver.get('https://s56.asp.srv.br/saude.pm.barretos.sp/login')
    
    driver.find_element('xpath', '//*[@id="vASPUSUARIOLOGINSENHA"]').send_keys(password)
    time.sleep(1)
    driver.find_element('xpath', '//*[@id="vASPUSUARIOLOGINLOGIN"]').send_keys(user)
    
    driver.find_element('id', 'IMGASPLOGIN').click()
    time.sleep(1)
    driver.find_element('xpath', '//*[@id="ext-gen9"]').click() #click menu
    time.sleep(1)
    driver.find_element('xpath', '//*[@id="ext-gen93"]').click() #click atendimento
    time.sleep(1)
    driver.find_element('xpath', '//*[@id="ext-gen150"]').click() #click Controle de Agravos
    time.sleep(1)
    
    remove_old_files(default_download, ".xls")
    download_agravos(CID_AGRAVOS)    
    concat_agravos_dengue_assessor(default_download)
    os.startfile("exportacoes_teste", "open")

    
if __name__ == "__main__":
    _agravos_dengue_assessor()
    print("Agravos de dengue do assessor concatenados na pasta exportacoes")






