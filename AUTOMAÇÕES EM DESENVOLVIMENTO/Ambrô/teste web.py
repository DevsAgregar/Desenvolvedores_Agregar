from config.settings import LOGIN, SENHA
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

def executar_script():
    # Diretório
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

    # Caminho do chromedriver
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    service = Service(chromedriver_path)

    # Inicializa o navegador
    navegador = webdriver.Chrome(service=service, options=chrome_options)

    navegador.get('https://accounts.tiny.com.br/realms/tiny/protocol/openid-connect/auth?client_id=tiny-webapp&redirect_uri=https://erp.tiny.com.br/login&scope=openid&response_type=code')
    time.sleep(5)

