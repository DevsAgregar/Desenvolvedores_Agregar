from config.settings import EMAIL_AMBRO_JR, SENHA_AMBRO_JR, EMAIL_SHOPEE, SENHA_SHOPEE
import os
import shutil
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time
import sys
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime

def executar_script():
    download_dir = "C:\\Users\\User\\Downloads"
    
    nome_perfil = 'Default'
    
    caminho_perfil = 'C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data'

    # Configs do chrome
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument(f'user-data-dir={caminho_perfil}')
    chrome_options.add_argument(f'profile-directory={nome_perfil}')

    # Instala o ChromeDriver
    chrome_install = ChromeDriverManager().install()

    # Caminho do ChromeDriver
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    service = Service(chromedriver_path)

    # Inicializa o navegador
    navegador = webdriver.Chrome(service=service, options=chrome_options)

    # Variável universal para usar ao decorrer do código
    data_atual = datetime.now().strftime("%d/%m/%Y")
    
    # Lista com os logins para iteração em diferentes acessos
    logins = [
        (EMAIL_AMBRO_JR, SENHA_AMBRO_JR),
        (EMAIL_SHOPEE, SENHA_SHOPEE)
    ]

    def login_tiny(email, senha):
        # Entra no tiny
        navegador.get('https://accounts.tiny.com.br/realms/tiny/protocol/openid-connect/auth?client_id=tiny-webapp&redirect_uri=https://erp.tiny.com.br/login&scope=openid&response_type=code')
        time.sleep(5)
        
        # Tenta clicar no botão que aparece quando alguém já está logado
        try:
            navegador.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]').click()
            time.sleep(5)
        except Exception as e:
            print('Botão não disponível')
            
        # Insere os dados de login e aperta em entrar
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input').send_keys(email)
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input').send_keys(senha)
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button').click()
        time.sleep(5)
        
        # Tenta clicar em outro botão de login que aparece após entrar no acesso
        try:
            navegador.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]').click()
            time.sleep(5)
        except Exception as e:
            print('Botão não disponível')

        time.sleep(5)
        
        
    def baixar_vendas_markup():
        def vendas_markup_ambrojr():
            # Acessa a aba "Análise de Markup"
            navegador.get('https://erp.tiny.com.br/relatorios_personalizados#/view/80')
            time.sleep(7)
            
            # Acessa o filtro de data
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/a').click()
            # Seleciona "Período"
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/div/div[2]/div/div[5]/button').click()
            time.sleep(2)
            # Insere as datas
            bt_dt_ini = navegador.find_element(By. XPATH, '/html/body/div[6]/div/div[2]/div/div[1]/div[4]/ul/li[1]/div/div[3]/div[1]/div/input')
            bt_dt_ini.clear()
            bt_dt_ini.send_keys('01/03/2024')
            
            bt_dt_fim = navegador.find_element(By.XPATH, '/html/body/div[6]/div/div[2]/div/div[1]/div[4]/ul/li[1]/div/div[3]/div[2]/div/input')
            bt_dt_fim.clear()
            bt_dt_fim.send_keys(data_atual)
            
            # Aplica o filtro de Período
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[1]/div/div[4]/button[1]').click()
            time.sleep(5)
            
            # Seleciona as situações
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/a').click()
            # Abre as opções
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/div/div[1]/div/button').click()
            # Desseleciona "Em aberto"
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/div/div[1]/div/ul/li[2]/a/div/label').click
            # Desseleciona "Cancelado"
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/div/div[1]/div/ul/li[11]/a/div/label').click()
            # Fecha as opções
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/div/div[1]/div/button').click()
            # Aplica o filtro das situações
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[4]/ul/li[2]/div/div[2]/button[1]').click()
            time.sleep(5)
            
            # Aperta no botão "download"
            navegador.find_element(By.XPATH, '//*[@id="root-relatorios-personalizados"]/div/div[1]/div[1]/div/div/div/button[2]').click()
            navegador.find_element(By.XPATH, '//*[@id="bs-modal"]/div/div/div/div[3]/button[2]').click()
executar_script()
    
        

