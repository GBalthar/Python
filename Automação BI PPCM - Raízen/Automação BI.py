#Este programa é uma cortesia de Guilherme Balthar.
#Data do projeto: 19/04/2023.
#Objetivo: Automação do processo de download das planilhas do TrakSYS para atualização do BI do PPCM.


#Inicialização
from selenium import webdriver #Importando a biblioteca de automação 'Selenium'
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time #Importando ferramenta de controle de tempo
import os #Importando ferramenta de renomear arquivos
print('\nSejam bem-vindos ao programa de automação do BI do PPCM!')
time.sleep(1)
print('\nCertifique-se de que sua conexão com a internet é de cabo LAN.\n')
time.sleep(1)

#Input do CS
cs = input('\nInsira seu CS no formato CSXXXXXX: ')
print('\n\nCS fornecido: ', cs)
time.sleep(1)
ok_cs = int(input('\nO CS fornecido está correto? \n\nDigite 0 para Sim\n\nDigite 1 para Não\n\nSua resposta: '))
while(ok_cs == 1):
    cs = input('\n\nInsira seu CS novamente: ')
    print('\n\nNova CS fornecido: ', cs)
    ok_cs = int(input('\nO CS agora está correto? \n\nDigite 0 para Sim\n\nDigite 1 para Não\n\nSua resposta: '))

#Input da data do dia atual
dia = input('\n\nInsira a data de hoje no formato DD/MM/AAAA: ')
print('\n\nData fornecida: ', dia)
time.sleep(1)
ok_dia = int(input('\nA data fornecida está correta? \n\nDigite 0 para Sim\n\nDigite 1 para Não\n\nSua resposta: '))
while(ok_dia == 1):
    dia = input('\n\nInsira a data de hoje novamente: ')
    print('\n\nNova data fornecida: ', dia)
    ok_dia = int(input('\nA data agora está correta? \n\nDigite 0 para Sim\n\nDigite 1 para Não\n\nSua resposta: '))


#Abrindo o TrakSYS
print('\n\nAbrindo o TrakSYS')
navegador_service = Service(executable_path='chromedriver.exe')
navegador = webdriver.Chrome(service=navegador_service)
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
navegador = webdriver.Chrome(options=options)
navegador.maximize_window()
navegador.get('http://aztraksysprd01/TS/Account/LogOn.aspx?ts_deny=true&ts_rurl=%2fTS%2fdefault.aspx')
time.sleep(25)


#Logando com Windows Login
navegador.find_element(by=By.XPATH, value= '//*[@id="loginTabs"]/li[2]/a').click()
time.sleep(2)
navegador.find_element(by=By.XPATH, value= '//*[@id="LoginWindowsButton"]').click()
time.sleep(20)


#Download OCME 1
#Aba KPI
navegador.find_element(by=By.XPATH, value= '//*[@id="formData"]/div[3]/section[1]/div[2]/ul/li[2]/a').click()
time.sleep(10)
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select OCME 1
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/ul/li[2]/a').click()
time.sleep(2)
#Filtro de grupo de tempo
select = Select(navegador.find_element(by=By.ID, value='FilterGroupBy_Filter'))
time.sleep(2)
#Selecionar dia
select.select_by_value('2')
time.sleep(2)
#Data inicial
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterDatePicker_Start"]').clear()
time.sleep(2)
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterDatePicker_Start"]').send_keys('01/01/2021')
time.sleep(2)
#Data final
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterDatePicker_End"]').clear()
time.sleep(2)
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterDatePicker_End"]').send_keys(dia)
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da OCME 1 realizado com sucesso.')


#Download OCME 2
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select OCME 2
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/ul/li[3]/a').click()
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da OCME 2 realizado com sucesso.')


#Download OCME 4
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select OCME 4
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/ul/li[4]/a').click()
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da OCME 4 realizado com sucesso.')


#Download Tambor 1-2
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select Tambor 1-2
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/ul/li[6]/a').click()
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da Tambor 1-2 realizado com sucesso.')


#Download Tambor 3-4
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select Tambor 3-4
button = navegador.find_element(by=By.XPATH, value='//*[@id="FilterSystem"]/div/div/div/ul/li[7]/a')
navegador.execute_script("arguments[0].click();", button)
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da Tambor 3-4 realizado com sucesso.')


#Download Tambor 5-6
#Filtro de equipamento
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/button/span[1]').click()
time.sleep(2)
#Unsellect all
navegador.find_element(by=By.XPATH, value= '//*[@id="FilterSystem"]/div/div/div/div[2]/div/button[2]').click()
time.sleep(2)
#Select Tambor 5-6
button = navegador.find_element(by=By.XPATH, value='//*[@id="FilterSystem"]/div/div/div/ul/li[8]/a')
navegador.execute_script("arguments[0].click();", button)
time.sleep(2)
#Refresh
button = navegador.find_element(by=By.XPATH, value='//*[@id="ButtonRefresh_Button"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(10)
button = navegador.find_element(by=By.XPATH, value='//*[@id="contentPage_MenuButtons_tsbtnitem_1"]')
navegador.execute_script("arguments[0].click();", button)
time.sleep(5)
print('\nDownload do OEE da Tambor 5-6 realizado com sucesso.')


#Renomeando arquivos baixados
old_ocme1 = rf"C:\Users\{cs}\Downloads\OEE.xlsx"
new_ocme1 = rf"C:\Users\{cs}\Downloads\OCME 1 OEE ANUAL.xlsx"
os.rename(old_ocme1, new_ocme1)
old_ocme2 = rf"C:\Users\{cs}\Downloads\OEE (1).xlsx"
new_ocme2 = rf"C:\Users\{cs}\Downloads\OCME 2 OEE ANUAL.xlsx"
os.rename(old_ocme2, new_ocme2)
old_ocme4 = rf"C:\Users\{cs}\Downloads\OEE (2).xlsx"
new_ocme4 = rf"C:\Users\{cs}\Downloads\OCME 4 OEE ANUAL.xlsx"
os.rename(old_ocme4, new_ocme4)
old_tambor12 = rf"C:\Users\{cs}\Downloads\OEE (3).xlsx"
new_tambor12 = rf"C:\Users\{cs}\Downloads\TAMBOR 1-2 OEE ANUAL.xlsx"
os.rename(old_tambor12, new_tambor12)
old_tambor34 = rf"C:\Users\{cs}\Downloads\OEE (4).xlsx"
new_tambor34 = rf"C:\Users\{cs}\Downloads\TAMBOR 3-4 OEE ANUAL.xlsx"
os.rename(old_tambor34, new_tambor34)
old_tambor56 = rf"C:\Users\{cs}\Downloads\OEE (5).xlsx"
new_tambor56 = rf"C:\Users\{cs}\Downloads\TAMBOR 5-6 OEE ANUAL.xlsx"
os.rename(old_tambor56, new_tambor56)


#Finalização
print('\n\nAgora basta fazer o upload dos arquivos para o SharePoint e clicar no botão atualizar no Power BI.\n\n\nObrigado por fazer uso desta ferramenta, seu processo já foi concluído!\n\n\nEste programa é uma cortesia de Guilherme Balthar.\n\n')
