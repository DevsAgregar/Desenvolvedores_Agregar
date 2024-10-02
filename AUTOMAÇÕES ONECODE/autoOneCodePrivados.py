import os
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.keys import Keys
from datetime import datetime

def executar_script():

    download_dir = "C:\\Users\\User\\Downloads"

    # Configurações do Chrome
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_options.add_argument('--start-maximized')

    # Instala o ChromeDriver
    chrome_install = ChromeDriverManager().install()

    # Caminho do chromedriver
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    service = Service(chromedriver_path)

    # Inicia o navegador
    navegador = webdriver.Chrome(service=service, options=chrome_options)

    # Abre o onecode
    navegador.get('https://agregar.onecode.chat/tickets')
    time.sleep(5)

    # Faz o login
    navegador.find_element(By.XPATH, '//*[@id="email"]').send_keys('pedrothiagodossantos74@gmail.com')
    navegador.find_element(By.XPATH, '//*[@id="password"]').send_keys('12345678')
    time.sleep(5)
    navegador.find_element(By.XPATH, '//*[@id="root"]/main/div/div[2]/div[1]/form/button').click()
    time.sleep(15)

    dia_da_semana = datetime.now().weekday()  # 0 = segunda, 4 = sexta
    clientes = []

    if dia_da_semana == 0:  # Segunda-feira
        clientes = []
        
    elif dia_da_semana == 1:  # Terça-feira
        clientes = []

    elif dia_da_semana == 2:  # Quarta-feira
        clientes = [
            {"nome": "OLGA TARTARI MEYER", "mensagem": "Olá Cleusa, muito bom dia! Segue o link de acesso do seu relatório financeiro: https://app.powerbi.com/view?r=eyJrIjoiODU5ZjEzN2ItZWFjYS00MWEyLWJlYmQtMWU4NWMxOWJkMDg1IiwidCI6ImNmZmVjMmI5LWQ1ODUtNGE2Yi04MzllLTA1MzVlNTljZjViZCJ9&pageName=ReportSection"}
            ]

    elif dia_da_semana == 3:  # Quinta-feira
        clientes = []

    elif dia_da_semana == 4:  # Sexta-feira
        clientes = []

    else:
        clientes = []

    if not clientes:
        navegador.quit()
        return

    for cliente in clientes:
        # Abre o One Code
        navegador.get('https://agregar.onecode.chat/tickets')
        time.sleep(5)

        # Pesquisa o cliente
        navegador.find_element(By.XPATH, '/html/body/div/div[2]/main/div/div/div[1]/div/div[1]/div/div[2]/div/div/div/button[4]').click()
        time.sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[3]/div/div[1]/div[1]/input').clear()
        navegador.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[3]/div/div[1]/div[1]/input').send_keys(cliente["nome"])
        time.sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[4]/div[11]/button').click()
        time.sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="simple-tabpanel-search"]/div/div/ul/div[1]/div').click()
        time.sleep(5)

        # Abre a conversa
        navegador.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[1]/div[2]/div/button[2]').click()
        time.sleep(5)

        # Envia a mensagem                
        navegador.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[4]/div[3]/div/div[1]/textarea[1]').clear()
        navegador.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[4]/div[3]/div/div[1]/textarea[1]').send_keys(cliente["mensagem"])
        time.sleep(5)  
        navegador.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[4]/div[3]/div/div[1]/textarea[1]').send_keys(Keys.ENTER)                   
        time.sleep(2)

        # Finaliza o ticket               
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[2]/main/div/div/div[1]/div/div[2]/div/div/div[1]/div[2]/div/button[4]').click()
        time.sleep(5)
        navegador.find_element(By.XPATH, '/html/body/div[3]/div[3]/div/div[2]/button[2]').click()
        time.sleep(10)

    # Fecha o navegador
    navegador.quit()

executar_script()