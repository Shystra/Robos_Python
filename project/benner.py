import undetected_chromedriver as webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pandas as pd
import gspread
import datetime

def baixar_relatrio():
    try:
        print('1    Iniciando Processo de atualização do CONSULT006 (DEMITIDOS & AFASTADOS)')
        dir = r"C:/Users/localuser/Documents/joao/Planilhas/TEMP PLAN BOT"
        dir2 = "C:/Users/localuser/Documents/joao/Planilhas/TEMP PLAN BOT/"
        print('2    Limpando os dados da pasta TEMP PLAN BOT')
        #Limpar Pasta antes de baixar
        for f in os.listdir(dir2):
            os.remove(os.path.join(dir2, f))

        #Processo de Manipular o navegador e Gerar Relatorio
        option = webdriver.ChromeOptions()
        #profile = "C:\\Users\\localuser\\AppData\\Local\\Google\\Chrome\\User Data\\Default"
        profile = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        option.add_argument(f"user-data-dir={profile}")
        #option.headless = True
        print('3    Login Realizado')
        driver = webdriver.Chrome(options=option,use_subprocess=True)
        driver.get("https://intersept.bennercloud.com.br/RH/Login")
        time.sleep(2)
        driver.find_element(By.ID , 'LoginButton').click()
        time.sleep(2)
        driver.get("https://intersept.bennercloud.com.br/RH/aga/a/modulos/relatorios.aspx?i=R_RELATORIOS_GRID&m=MAIN")
        time.sleep(2)
        driver.find_element(By.ID , 'ctl00_Main_RELATRIOS_FilterControl_GERAL_1__CODIGO').clear()
        time.sleep(1)
        print('4    Acessando o CONSULT006')
        driver.find_element(By.ID , 'ctl00_Main_RELATRIOS_FilterControl_GERAL_1__CODIGO').send_keys('CONSULT006')
        time.sleep(10)
        driver.find_element(By.ID , 'ctl00_Main_RELATRIOS_FilterControl_FilterButton').click()
        time.sleep(3)
        driver.find_elements(By.CLASS_NAME , 'fa-up-from-line')[0].click()
        WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.ID , 'CMD_EXPORTARXLS'))).click()
        time.sleep(3)
        #driver.find_elements(By.CLASS_NAME , 'no-js')[0].click()
        print('Modal Aberto')
        iframe = driver.find_element(By.XPATH, "//div[@class='modal-body']//iframe")
        driver.switch_to.frame(iframe)
        time.sleep(2)
        driver.find_elements(By.CLASS_NAME, 'select2-search__field')[5].send_keys(' ')
        time.sleep(2)
        driver.find_elements(By.CLASS_NAME, 'select2-results__option')[2].click()
        time.sleep(2)
        #action = ActionChains(driver)
        #action.key_down(Keys.CONTROL).send_keys(Keys.ENTER).perform()
        time.sleep(2)
        driver.find_element(By.CLASS_NAME, 'btn-save').click()
        driver.switch_to.default_content()
        print('Modal Fechado')
        time.sleep(5)
        driver.find_elements(By.CLASS_NAME , 'dropdown-toggle')[0].click()
        print('5    Aguardando a planilha ser gerada pela plataforma da Benner')
        WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.CLASS_NAME, 'finished')))
        time.sleep(2)
        driver.find_elements(By.CLASS_NAME , 'finished')[0].click()
        print('6    Planilha baixada')
        time.sleep(35)
        print('7    Navegador baixado')

        print('8    Tratando os Dados da planilha')
        # Abrir Planilha e selecionar colunas
        df_pd = pd.read_excel(dir2 + 'Relatório+Relatorio+Informações+Funcionais+Folha.xlsx', engine="openpyxl")[[
                                                                                                                      'Nome'
                                                                                                                     ,'Unidade'
                                                                                                                     ,'Demissão'
                                                                                                                     ,'Matrícula'
                                                                                                                     ,'Função'
                                                                                                                     ,'Desc. Situação'
                                                                                                                ]]

        df_pd.loc[df_pd['Nome'] == 'Nome']

        # Formatar os dados da Planilha
        cliente = []
        for row in df_pd.to_numpy():
            cliente.append([str(elem) for elem in row])

        m = [row for row in cliente if 'nan' != row[1] if 'Unidade' != row[1]]
        print('9    ACESSANDO A API DA GOOGLE SHEETS PARA REALIZAR O UPLOAD')
        # API/KEY: GOOGLE SHEETS  - IMPORTANTE: compartilhar a G-SHEETS com o E-mail: python-connect-sheets@pythonsheets-344316.iam.gserviceaccount.com
        CODE = 'CREDENCIAL DE DESTINO'
        DICT = {}

        gc = gspread.service_account(filename='key.json')
        sh = gc.open_by_key(CODE)
        ws = sh.worksheet('CONSULT006')
        # row = len(ws.get_all_records())+2
        # print(x)
        print('10    LIMPANDAO A PLANILHA')
        #Limpando Planilha
        sh.values_clear("CONSULT006!A2:F90000")

        # Salvando os dados na G-SHEETS
        sh.values_update(
            'CONSULT006!A2',
            params={
                'valueInputOption': 'RAW'
            },
            body={
                'values': m
            }
        )

        CODE = '1Zqgv6UMUw9NKo3N_9PfnCEs9oe8AQw52QxPPYO-Qdjo'
        DICT = {}

        gc = gspread.service_account(filename='key.json')
        sh = gc.open_by_key(CODE)
        ws = sh.worksheet('HIST')
        # row = len(ws.get_all_records())+2
        # print(x)

        # Salvando os dados na G-SHEETS
        sh.values_update(
            'HIST!B4',
            params={
                'valueInputOption': 'RAW'
            },
            body={
                'values': [[str( datetime.datetime.now() )]]
            }
        )

        print('10    UPLOAD REALIZADO COM SUCESSO')
    except FileNotFoundError as error:
        print("Erro: " + error)
        resposta = input('Gostaria de tentar baixar novamente? (Y/n)')
        if  resposta.upper() == 'Y':
            print('REINICIANDO...')
            driver.quit()
            baixar_relatrio()

        else:
            print('FIM...')

baixar_relatrio()
