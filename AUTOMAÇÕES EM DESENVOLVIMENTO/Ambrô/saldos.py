from config.settings import EMAIL_AMBRO_JR, SENHA_AMBRO_JR
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import load_workbook
import os

# Diretório de download temporário
download_dir = "C:\\Users\\User\\Downloads"

# Configurações do Chrome para definir o diretório de download
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
chrome_options.add_argument("--start-maximized")

# Instala o ChromeDriver
chrome_install = ChromeDriverManager().install()

# Caminho do chromedriver
folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")
service = Service(chromedriver_path)

# Inicializa o navegador
navegador = webdriver.Chrome(service=service, options=chrome_options)


def login_tiny():
    # Acessa o tiny
    navegador.get('https://accounts.tiny.com.br/realms/tiny/protocol/openid-connect/auth?client_id=tiny-webapp&redirect_uri=https://erp.tiny.com.br/login&scope=openid&response_type=code')

    # Faz o login
    navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input').send_keys(EMAIL_AMBRO_JR)
    navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input').send_keys(SENHA_AMBRO_JR)
    navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button').click()
    time.sleep(10)

    # Faz o login se outra conta está logada
    elemento = navegador.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]')
    actions = ActionChains(navegador)
    actions.move_to_element(elemento).click().perform()
    time.sleep(10)


data_atual = datetime.now()


def atualizar_saldos():
    try:
        # Navega até a página desejada
        navegador.get('https://erp.tiny.com.br/caixa')
        time.sleep(5)

        # Clicar nos elementos e navegar conforme necessário
        navegador.find_element(By.XPATH, '/html/body/div[6]/div/div[4]/div[1]/div[3]/ul/li[3]/a').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="page-wrapper"]/div[4]/div[2]/div[2]/ul/li/a').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="item-conta-todas"]').click()
        time.sleep(6)

        # Dicionário de XPaths e células correspondentes na planilha
        xpath_celula_map = {
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[1]/div/span': 'B2',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[2]/div/span': 'B3',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[3]/div/span': 'B4',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[4]/div/span': 'B5',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[5]/div/span': 'B6',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[6]/div/span': 'B7',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[7]/div/span': 'B8',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[8]/div/span': 'B9',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[9]/div/span': 'B10',
            '/html/body/div[6]/div/div[4]/div[2]/div[2]/div/span/div[10]/div/span': 'B11'
        }

        # Carregar a planilha existente
        wb = load_workbook('G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\FINANCEIRO\\SALDO DAS CONTAS.xlsx')

        for xpath, celula in xpath_celula_map.items():
            try:
                # Encontrar o elemento com base no XPath e extrair o texto
                elemento = navegador.find_element(By.XPATH, xpath)
                valor = elemento.text.strip()

                # Selecionar a planilha e escrever o valor na célula especificada
                sheet = wb.active  # ou selecione a planilha desejada: wb['Nome_da_Planilha']
                sheet[celula] = valor

            except Exception as e:
                print(f"Erro ao extrair valor do XPath {xpath}:", e)

        # Salvar a planilha de volta no mesmo arquivo
        wb.save('G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\FINANCEIRO\\SALDO DAS CONTAS.xlsx')

    except Exception as e:
        print("Erro durante a execução do script:", e)

    finally:
        # Fechar o navegador ao finalizar
        navegador.quit()
        time.sleep(5)



login_tiny()
atualizar_saldos()