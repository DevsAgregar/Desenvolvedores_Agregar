import os
import shutil
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
import sys

def executar_script():
    max_tentativas = 2
    tentativas = 0
    
    while tentativas < max_tentativas:
        navegador = None
        try:
            # Diretório de download
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
            
            def login_saib_erp():
                # Acessa o Saib ERP
                navegador.get('https://erp.saibweb.com.br/auth')
                time.sleep(2)
                
                # Preenche os campos de Usuário e Senha e faz login
                navegador.find_element(By.XPATH, '//*[@id="username"]').send_keys('BRAZZICOR.AGREGAR')
                time.sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="password"]').send_keys('123456')
                time.sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div/div[2]/form/button').click()
                time.sleep(5)
                
            def baixar_relatorios():
                
                for relatorio, destino in relatorios_destinos.items():
                    # Acessa a aba de scripts
                    navegador.get('https://erp.saibweb.com.br/admin/exportador/scripts/execucao')
                    time.sleep(5)
                    
                    # Faz todo o processo para gerar o relatório
                    escolher_modulo = navegador.find_element(By.XPATH, '//*[@id="scrollable-force-tabpanel-0"]/div/form/div/div[2]/div/div')
                    escolher_modulo.click()
                    pyautogui.write('FINANCEIRO')
                    pyautogui.press('enter')
                    time.sleep(2)
                    
                    escolher_script = navegador.find_element(By.XPATH, '//*[@id="scrollable-force-tabpanel-0"]/div/form/div[2]/div[2]/div/div/div[1]/div[2]')
                    escolher_script.click()
                    pyautogui.write(relatorio)   
                    pyautogui.press('enter')
                    time.sleep(2)
                    
                    # Aperta na Lupa e passa para a próxima etapa
                    navegador.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div[2]/div/div[2]/div[2]/div/button').click()
                    time.sleep(2)
                    
                    # Filtra a data                                 
                    data_inicial = navegador.find_element(By.XPATH, '//*[@id="scrollable-force-tabpanel-1"]/div/form/div[2]/div[2]/div/div/input')
                    data_inicial.click()
                    pyautogui.hotkey('ctrl', 'a')
                    data_inicial.send_keys('01/07/2024')
                    time.sleep(2)
                    
                    # Gera o relatório
                    navegador.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div[2]/div/div[2]/div[2]/div/button').click()
                    time.sleep(10)
                    
                    # Baixa o relatório
                    navegador.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div[2]/div/div[2]/div[2]/div/div/button[3]').click()
                    time.sleep(5)
                    
                    
                    # Caminho do arquivo baixado e do arquivo de destino
                    downloaded_file = os.path.join(download_dir, 'Script.xlsx')
                    destination_file = destino
                    
                    # Move o arquivo
                    if os.path.exists(downloaded_file):
                        shutil.move(downloaded_file, destination_file)
                    
            relatorios_destinos = {
                "96 - MOV CX BCO": "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BRAZZICOR TINTAS\\3. Finanças\\4- Projeto BI\\01. Banco de Dados\\02 - FINANCEIRO\\EXTRATO FINANCEIRO a partir do mês 08.xlsx",
                "88 - CONTAS A RECEBER - POR VENCIMENTO": "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BRAZZICOR TINTAS\\3. Finanças\\4- Projeto BI\\01. Banco de Dados\\02 - FINANCEIRO\\CONTAS A RECEBER E A PAGAR\\CONTAS A RECEBER FUTURO.xlsx",
                "116 - CONTAS A PAGAR COM CENTRO DE CUSTOS": "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BRAZZICOR TINTAS\\3. Finanças\\4- Projeto BI\\01. Banco de Dados\\02 - FINANCEIRO\\CONTAS A RECEBER E A PAGAR\\CONTAS A PAGAR FUTURO.xlsx"
            }

                    
            login_saib_erp()
            baixar_relatorios()
                
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
