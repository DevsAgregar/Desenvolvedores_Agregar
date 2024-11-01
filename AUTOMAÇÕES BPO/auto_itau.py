from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# Diretório
download_dir = "C:\\Users\\User\\Downloads"

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

# Instala o ChromeDriver
chrome_install = ChromeDriverManager().install()

# Caminho do chromedriver
folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")
service = Service(chromedriver_path)

# Inicializa o navegador
navegador = webdriver.Chrome(service=service, options=chrome_options)

navegador.get('https://www.itau.com.br/empresas')


# Lista de XPaths dos botões
xpaths_botoes = [
    '',
    '',
    '',
    '',
    '',
    ''
]

# Função para encontrar e clicar no botão com o dígito correto
def clicar_botao_com_digito(digito):
    for xpath in xpaths_botoes:
        botao = navegador.find_element(By.XPATH, xpath)
        if digito in botao.text:
            botao.click()
            break

# Sua senha
senha = "123456"

# Loop para clicar nos botões na ordem da senha
for digito in senha:
    clicar_botao_com_digito(digito)
    time.sleep(1)  # Pequena pausa para garantir que o clique foi registrado

# Fechar o navegador
time.sleep(2)  # Espera para ver o resultado
navegador.quit()