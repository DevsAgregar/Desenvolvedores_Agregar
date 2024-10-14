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
    max_tentativas = 2
    tentativas = 0
    
    while tentativas < max_tentativas:
        navegador = None
        try:
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

            # Abre o Conta Azul
            navegador.get('https://mais.contaazul.com/#/login')
            time.sleep(5)

            # Login no Conta Azul
            navegador.find_element(
                By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div/div/div[2]/form/div/div/div[1]/div/div/div/input').send_keys('pedrothiagoagregar@gmail.com')
            time.sleep(2)
            navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div/div/div[2]/form/div/div/div[2]/div/div/div/div/div/input').send_keys('Pedroth45.')
            time.sleep(2)
            navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div/div/div[2]/form/div/div/div[3]/div[1]/div/span/button').click()
            time.sleep(10)

            # Fecha pop-up 
            pyautogui.press('esc')
            time.sleep(2)

            # Pesquisa o Cliente
            navegador.find_element(
                By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[3]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div[1]/div[1]/div/div/form/div/div/div/div[1]/input').send_keys('Alves E Lima Carretas E Traillers Ltda')
            time.sleep(2)
            pyautogui.hotkey('enter')
            time.sleep(2)

            # Seleciona o cliente
            navegador.find_element(
                By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[3]/div[2]/div/div/div[3]/div/div/div/div/div/div[2]/div/div/div/div/table/tbody/tr/td[1]').click()
            time.sleep(2)

            # Abre o CA Pro
            navegador.find_element(
                By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/aside/div/div/div/div/ul[2]/li[2]/div[2]/ul/li[2]/a/div/div').click()
            time.sleep(10)

            # Muda o foco para a nova aba
            navegador.switch_to.window(navegador.window_handles[-1])

            # Abre o extrato
            navegador.get('https://app.contaazul.com/#/ca/financeiro/extrato')
            time.sleep(3)

            # Espera até que o botão de período esteja clicável e clica
            wait = WebDriverWait(navegador, 10)
            period_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="gateway"]/section/div[3]/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[1]/div/div/div/div/div/div[1]/span/button')))
            period_button.click()
            time.sleep(2)

            # Espera até que o filtro de todo o período esteja clicável e clica
            all_period_filter = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="gateway"]/section/div[3]/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[1]/div/div/div/div/div/div[2]/div[7]')))
            all_period_filter.click()
            time.sleep(10)

            # Espera até que o botão de exportar relatório esteja clicável e clica
            export_button = wait.until(EC.element_to_be_clickable(  
                (By.XPATH, '//*[@id="gateway"]/section/div[2]/div/nav/div/div/div[1]/div[2]/div/div[2]/span/button')))
            export_button.click()

            # Espera o download ser concluído
            time.sleep(15)

            # Caminho do arquivo baixado
            downloaded_file = os.path.join(download_dir, "extrato_financeiro.xls")

            # Caminho do arquivo de destino
            destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar-Clientes Inativos\\AL CARRETAS\\3. Finanças\\5 - Power Bi\\01. BANCO DE DADOS\\extrato_financeiro.xls"

            # Move o arquivo baixado para o diretório de destino, substituindo o arquivo existente
            if os.path.exists(downloaded_file):
                shutil.move(downloaded_file, destination_file)

            return 0 # 0 para sucesso
        
        except Exception as e:
            tentativas += 1
            if tentativas < max_tentativas:
                time.sleep(5)
            else: 
                return 1 # 1 para falha

        finally:
            if navegador:
                navegador.quit()
                
    return 1


sys.exit(executar_script())