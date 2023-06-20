import undetected_chromedriver as webdriver
import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
import os
from zipfile import ZipFile
import pandas as pd
import gspread
import datetime

dir = r"C:/Users/localuser/Documents/joao/Planilhas/TEMP PLAN BOT"
dir2 = "C:/Users/localuser/Documents/joao/Planilhas/TEMP PLAN BOT/"

#Limpar Pasta antes de baixar
print("1    Iniciando Processo de atualização do R05")
print("2    Limpando os dados da pasta TEMP PLAN BOT")
for f in os.listdir(dir2):
    os.remove(os.path.join(dir2, f))

#Processo de Manipular o navegador
option = webdriver.ChromeOptions()
#profile = "C:\\Users\\localuser\\AppData\\Local\\Google\\Chrome\\User Data\\Default"
profile = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
option.add_argument(f"user-data-dir={profile}")
#option.headless = True
driver = webdriver.Chrome(options=option,use_subprocess=True)
driver.get("https://gestao.pontotel.com.br/#/cognito/login/")
time.sleep(2)
# matriculas = ['4939','4539', '0279']
#time.sleep(60*(random.randint(1, 1)))
#print("Teste")
#https://pt.stackoverflow.com/questions/549127/checar-a-exist%C3%AAncia-de-um-elemento-usando-python-selenium
# for mat in matriculas:
#     pergunta = input('Gostaria de bater o Ponto do: ' + mat + ' ?')
#     if pergunta == 'S':
#driver.find_element(By.CLASS_NAME, "p-textfield__input").send_keys('joaofragoso@intersept.com.br')
time.sleep(2)
driver.find_elements(By.CLASS_NAME , 'p-btn__content')[0].click()
time.sleep(2)
driver.find_elements(By.CLASS_NAME , 'p-btn__content')[1].click()
time.sleep(15)
print('3    Login Realizado')
driver.get("https://gestao.pontotel.com.br/#/fechamentos")
time.sleep(5)
print('4    Acessando o R05')
driver.find_element(By.XPATH, "//span[contains(., 'R05 - Planilha de apontamentos dos empregados no mês')]").click()
#driver.find_elements(By.CLASS_NAME , 'fs-15')[6].click()
time.sleep(5)
WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.CLASS_NAME , 'multiselect__placeholder'))).click()
WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.CLASS_NAME , 'multiselect__option'))).click()
time.sleep(2)
driver.find_elements(By.CLASS_NAME , 'multiselect__placeholder')[0].click()
time.sleep(2)
driver.find_element(By.XPATH, "//span[contains(., 'PIS, função e empregador dos empregados')]").click()
time.sleep(1)
driver.find_element(By.XPATH, "//span[contains(., 'supervisor do local de trabalho')]").click()
time.sleep(1)
driver.find_element(By.XPATH, "//span[contains(., 'estados e cidades dos locais de trabalho')]").click()
time.sleep(4)
driver.find_element(By.CLASS_NAME , 'p-btn__content').click()
time.sleep(5)
driver.find_element(By.CLASS_NAME , 'swal2-confirm').click()
time.sleep(3)
driver.find_element(By.CLASS_NAME , 'main-header__button--category').click()
time.sleep(3)
#Status_Rel = driver.find_elements(By.CLASS_NAME , "descricao")[0].get_attribute("innerHTML")
#relatórios em andamento (1)
# driver.find_elements(By.CLASS_NAME , 'btn3d')[3].click()
# time.sleep(2)
# driver.find_element(By.CLASS_NAME , "confirm").click()
# time.sleep(2)
#driver.get("https://www.youtube.com/watch?v=dQw4w9WgXcQ/")
# time.sleep(1000)
# driver.maximize_window()

while not os.listdir(dir) != []:
    time.sleep(5)
    #print(Status_Rel)
if os.listdir(dir) == []:
    print("Sem Arquivos")
else:
    print("Planilha Baixada")
    time.sleep(10)

    # Entrair Zip
    zip = os.listdir(path=dir)
    zip2 = [x for x in zip if '.zip' in x][0]

    # Desconpactar Zip
    #print(zip2)
    with ZipFile(dir2 + zip2, 'r') as file:
        file.extractall(dir2)

    # Buscar Planilha
    list_1 = os.listdir(path=dir)
    Plan = [x for x in list_1 if '.xlsx' in x][0]

    # Abrir Planilha e selecionar colunas
    df_pd = pd.read_excel(dir2 + Plan, engine="openpyxl")[['Código'
                                                            , 'Nome'
                                                            , 'Código'
                                                            , 'Senha'
                                                            , 'Data de Admissão'
                                                            , 'Local de Trabalho'
                                                            , 'Função'
                                                            , 'CNPJ'
                                                            , 'Status'
                                                            , 'CPF'
                                                            , 'RG'
                                                            , 'Supervisor Responsavel'
                                                            , 'Razão Social'
                                                            ,  'Estado'
                                                        ]]

    # Ajustar Nome das Empresas
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['LTDA INTERSEPT'], 'LTDA')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['Ivandir - Intersept Franchising'], 'IVANDIR')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['INTERSEPT VIGILÂNCIA JOINVILLE'], 'VGT')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['Intersept Vigilancia'], 'VGT')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['Intersept Vigilância - RS'], 'VGT')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['Madife'], 'MADIFE')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['Intersept Comercio'], 'COMERCIO')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['INTERSEPT HOLDING LTDA'], 'MULTISEG')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['IRIS BS SYSTEM'], 'IRIS')
    df_pd['Razão Social'] = df_pd['Razão Social'].replace(['INTERSAT RASTREAMENTO DE VEICULO LTDA'], 'INTERSAT')

    # Formartar os dados da planilha
    cliente = []
    for row in df_pd.to_numpy():  # Looping over returned rows and printing them
        # my_list = [elem for elem in row]
        cliente.append([str(elem) for elem in row])

    # API/KEY: GOOGLE SHEETS  - IMPORTANTE: compartilhar a G-SHEETS com o E-mail: python-connect-sheets@pythonsheets-344316.iam.gserviceaccount.com
    CODE = '1Zqgv6UMUw9NKo3N_9PfnCEs9oe8AQw52QxPPYO-Qdjo'
    DICT = {}

    gc = gspread.service_account(filename='key.json')
    sh = gc.open_by_key(CODE)
    ws = sh.worksheet('R05')
    # row = len(ws.get_all_records())+2
    # print(x)

    # Limpando Planilha
    sh.values_clear("R05!A2:M10000")

    # Salvando os dados na G-SHEETS
    # https://docs.google.com/spreadsheets/d/1Zqgv6UMUw9NKo3N_9PfnCEs9oe8AQw52QxPPYO-Qdjo/edit#gid=0
    sh.values_update(
        'R05!A2',
        params={
            'valueInputOption': 'RAW'
        },
        body={
            'values': cliente
        }
    )

    CODE = 'CREDENCIAL DE DESTINO'
    DICT = {}

    gc = gspread.service_account(filename='key.json')
    sh = gc.open_by_key(CODE)
    ws = sh.worksheet('HIST')
    # row = len(ws.get_all_records())+2
    # print(x)

    # Salvando os dados na G-SHEETS
    sh.values_update(
        'HIST!B2',
        params={
            'valueInputOption': 'RAW'
        },
        body={
            'values': [[str(datetime.datetime.now())]]
        }
    )