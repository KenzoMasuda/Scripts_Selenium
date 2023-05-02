# -*- coding: utf-8 -*-
"""
Created on Wed Feb  1 14:16:57 2023

@author: KENZO
"""

def _agravos_dengue_xi():
    
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import Select
    # from selenium.webdriver.support.ui import WebDriverWait
    import time
    from datetime import datetime
    from datetime import date
    # import itertools
    import os
    from glob import glob
    from sys import platform
    from  pathlib import Path
    import zipfile as zp
    #from os import chdir, getcwd, listdir
    from dbfread import DBF
    import chardet #detecta codificação de arquivos
    import pandas as pd
    
    '''def delete_old_export(path):
        list_export = []
        path += r"/*.dbf"
        list_export = glob(path)
        
        if len(list_export) == 0:
            print("Sem Exportações Antigas")
        else:
            for export in list_export:
                os.remove(export)
    '''  
    
    def get_name(caminho_completo):
        caminho = Path(caminho_completo)
        name = caminho.stem
        return name
    
    '''def lista_exportacoes(path): #pode ser melhorado, o glob já faz isso
        list_agravos = []
        chdir(path)
        getcwd()
        
        for file in listdir():
            if file.endswith('.dbf'):
                list_agravos.append(file)
        return list_agravos
    '''
       
    def remove_old_files(path, extension=None):
        for file in os.listdir(path):
            if extension:
                if Path(path, file).suffix == extension:
                    os.remove(Path(path, file))
            else:
                os.remove(Path(path, file))
                
    def corrige_arquivo(path):
        list_agravos_dbf = glob(path + "\*.dbf")
        list_agravos_basename = []
        df = []
        
        for agravos in list_agravos_dbf:
            list_agravos_basename.append(get_name(agravos))
            encoding = get_encoding_file(agravos)
            table = DBF(agravos, encoding = encoding)
            data = [dict(record) for record in table]
            df.append(pd.DataFrame(data))
        
                
        for i, agravo in enumerate(list_agravos_basename):
            if(df[i].columns[0] == "NU_NOTIFIC"):
                df[i].to_excel(path + "\\" + agravo + ".xlsx", index=False)
            else: 
                print("Corrigindo arquivo : " + agravo);
                index = df[i].columns.get_loc("NU_NOTIFIC")
                df_sup = df[i].iloc[:, :index] #pega o pedaço do dataframe fora de posição
                df[i] = df[i].drop(df_sup, axis = 1) #apaga aquele pedaço
                df[i] = pd.concat([df[i], df_sup], axis=1) #concatena o pedaço no final novamente
                df[i].to_excel(path + "\\" + agravo + ".xlsx", index = False)
    
    def concat_agravos_dengue_chikun(path):
        list_agravos = glob(path + "\*.xlsx")
        
        df = []
    
        for agravos in list_agravos:
            df.append(pd.read_excel(agravos))
            
        df_chikun = pd.concat([df[0], df[1]])
        df_chikun.to_excel(path + "\\" + "FebreChiSinan2.xlsx", index=False)
        df_dengue = pd.concat([df[2], df[3]])
        df_dengue.to_excel(path + "\\" + "positivos2.xlsx", index=False)
        
    def get_encoding_file(file): #problema para ler arquivos DBF, motivo: decodificação
        with open(file, 'rb') as f:
            raw_data = f.read()
            encoding = chardet.detect(raw_data)['encoding']
            return encoding
        #encoding dos arquivos DBF = ISO-8859-1
        
    def extrair_zip(path):
        os.startfile(path)
        list_path_exportacoes = []
        list_path_exportacoes = glob(os.path.join(path, '*.zip'))
        
        for i in list_path_exportacoes:
            #member_name = (zp.ZipFile(i, 'r').namelist()[0]) //quando necessário extrair somente 1 membro do .zip
            #zp.ZipFile(i, 'r').extract(member_name, path)
            zp.ZipFile(i, 'r').extractall(path)
            time.sleep(3)
            os.remove(i)
            
    def obter_num_export():
        cond = True #condição
        while cond != False:
            try:
                num_export = driver.find_element('xpath', '//*[@id="form:resultado"]/span').text
                tam = len(num_export)
                num_export_DBF.append(num_export[8:tam-1])
                cond = False
            except:
                time.sleep(1)
                continue
            
    def download_DBF(exports):
        cond = True #condição
        for i in exports:
            while cond != False:
                try:
                    xpath_link = f"//a[contains(@href, '{i}')]"
                    driver.find_element('xpath', xpath_link).click()
                    cond = False
                except:
                    time.sleep(4)
                    driver.refresh()
                    continue
            cond = True
    
    datas_inicio = ["01-01-2022", "01-01-2023"]
    datas_final = ["31-12-2022"]
    agravo = ["DENGUE", "FEBRE DE CHIKUNGUNYA"]
    num_export_DBF = []
    
    hoje = date.today().strftime("%d-%m-%Y")
    datas_final.append(hoje)
    
    user = os.environ.get('login_sinan')
    password = os.environ.get('pass_sinan')

    diretorio_atual = os.getcwd()
    caminho_pasta_download = os.path.join(diretorio_atual, "exportacoes")
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
    #1°Login + MENU solicitação exportação DBF
    driver.get('http://sinan.saude.gov.br/sinan/login/login.jsf')
    
    driver.find_element('id', 'form:username').send_keys(user)
    driver.find_element('id', 'form:password').send_keys(password)
    
    driver.find_element('xpath', '/html/body/div[4]/form/fieldset/div[4]/input').click()
    
    driver.find_element('id', 'barraMenu:j_id52_span').click()
    time.sleep(1)
    driver.find_element('id', 'barraMenu:j_id53:anchor').click()
    
    # --------------------------------------------------------------
    #2°Etapa: preencher formulário 
    for index in agravo:
        for i in range(2):
            driver.find_element('id', 'form:consulta_dataInicialInputDate').click()
            driver.find_element('id', 'form:consulta_dataInicialInputDate').send_keys(datas_inicio[i])
            #time.sleep(2)
            driver.find_element('id', 'form:consulta_dataFinalInputDate').click()
            driver.find_element('id', 'form:consulta_dataFinalInputDate').send_keys(datas_final[i])
            #time.sleep(2)
    
            Select(driver.find_element('id','form:tipoUf')).select_by_value('2')
            #time.sleep(2)
    
            Select(driver.find_element('xpath','//*[@id="form:formulario"]/fieldset/span[5]/select')).select_by_visible_text(index)
    
            driver.find_element('xpath', '//*[@id="form:formulario"]/input').click()
            driver.find_element('id','form:j_id128').click()
            obter_num_export()
            driver.refresh()
        
    driver.find_element('id', 'barraMenu:j_id52_span').click()
    time.sleep(1) 
    driver.find_element('id', 'barraMenu:j_id56:anchor').click()
    
    download_DBF(num_export_DBF)
    
    time.sleep(3)
    driver.close()
    remove_old_files(report_download_path, ".xlsx")
    extrair_zip(report_download_path)
    corrige_arquivo(report_download_path)
    remove_old_files(report_download_path, ".dbf")
    
    concat_agravos_dengue_chikun(report_download_path)
    
        
     
if __name__ == '__main__':
    _agravos_dengue_xi()
    print(r"Arquivos concatenados dentro da pasta 'exportacoes' do diretório atual")
