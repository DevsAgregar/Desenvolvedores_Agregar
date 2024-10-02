import os
from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
import pyautogui
import time
from selenium.webdriver.common.keys import Keys

def baixar_relatorio():
    # Configuração do Edge
    edge_options = EdgeOptions()
    edge_service = EdgeService(EdgeChromiumDriverManager().install())
    navegador_edge = webdriver.Edge(service=edge_service, options=edge_options)


    # Abre o link do BI
    navegador_edge.get('https://app.powerbi.com/groups/3bc0a327-0471-4ea3-bcbb-c8a90c77fd8c/reports/19dde639-a0b4-4e45-a1a1-b824910aa06c/ReportSection?experience=power-bi')
    time.sleep(10)

    # Clica no botão 'Período sem vender'
    navegador_edge.find_element(By.XPATH, '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/report/exploration-container/div/div/docking-container/div/div/exploration-fluent-navigation/section/nav/mat-action-list/button[10]/span/span[3]').click()
    time.sleep(2)

    # Seleciona a tabela
    navegador_edge.find_element(By.XPATH, '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/report/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[8]/transform/div/div[3]/div/div/div/div/div/div/h3').click()
    time.sleep(2)

    # Clica nos três pontinhos
    navegador_edge.find_element(By.XPATH, '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/report/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[8]/transform/div/visual-container-header/div/div/div/visual-container-options-menu/visual-header-item-container/div/button/i').click()
    time.sleep(2)

    # Clica em 'exportar dados'
    navegador_edge.find_element(By.XPATH, '/html/body/div[2]/div[4]/div/ng-component/pbi-menu/button[4]').click()
    time.sleep(2)

    # Clica em 'Exportar'
    navegador_edge.find_element(By.XPATH, '/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/export-data-dialog/mat-dialog-actions/button[1]').click()
    time.sleep(5)

    # Fecha o navegador Edge
    navegador_edge.quit()


def enviar_relatorio():
    
    # Diretório de download temporário
    download_dir = "C:\\Users\\User\\Downloads"

    # Configurações do Chrome para definir o diretório de download
    chrome_options = ChromeOptions()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    # Instala o ChromeDriver
    chrome_install = ChromeDriverManager().install()

    # Caminho do chromedriver
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    service = ChromeService(chromedriver_path)

    # Inicializa o navegador
    navegador_chrome = webdriver.Chrome(service=service, options=chrome_options)

    # Abre o onecode
    navegador_chrome.get('https://agregar.onecode.chat/tickets')

    # Faz o login
    navegador_chrome.find_element(By.XPATH, '//*[@id="email"]').send_keys('pedrothiagodossantos74@gmail.com')
    navegador_chrome.find_element(By.XPATH, '//*[@id="password"]').send_keys('12345678')
    time.sleep(5)
    navegador_chrome.find_element(By.XPATH, '//*[@id="root"]/main/div/div[2]/div[1]/form/button').click()
    time.sleep(5)

    # Pesquisa o cliente
    navegador_chrome.find_element(By.XPATH, '/html/body/div/div[2]/main/div/div/div[1]/div/div[1]/div/div[2]/div/div/div/button[4]').click()
    time.sleep(5)
    navegador_chrome.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[3]/div/div[1]/div[1]/input').clear()
    navegador_chrome.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[3]/div/div[1]/div[1]/input').send_keys('Testes Agregar')
    time.sleep(5)
    navegador_chrome.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/div[1]/div/div[1]/div/div[4]/div[11]/button').click()
    time.sleep(5)
    navegador_chrome.find_element(By.XPATH, '//*[@id="simple-tabpanel-search"]/div/div/ul/div[1]/div').click()
    time.sleep(5)

    # Abre a conversa
    navegador_chrome.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[1]/div[2]/div/button[2]').click()
    time.sleep(5)

    # Abre os documentos                     
    navegador_chrome.find_element(By.XPATH, '/html/body/div/div[2]/main/div/div/div[1]/div/div[2]/div/div/div[4]/div[3]/label/span').click()
    time.sleep(2)

    # Abre os downloads
    pyautogui.write('Downloads')
    pyautogui.hotkey('enter')
    time.sleep(1)

    # Seleciona o relatório
    pyautogui.write('DIAS SEM VENDER POR CLIENTE.xlsx')
    pyautogui.hotkey('enter')
    time.sleep(1)

    # Envia a mensagem
    pyautogui.write('Boa Tarde!! Segue a planilha de DIAS SEM VENDER POR CLIENTE')
    time.sleep(3)
    navegador_chrome.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[4]/div[3]/div/div[1]/textarea[1]').send_keys(Keys.ENTER)
    time.sleep(7)

    # Finaliza o ticket
    navegador_chrome.find_element(By.XPATH, '//*[@id="drawer-container"]/div/div[1]/div[2]/div/button[3]').click()
    time.sleep(5)
    navegador_chrome.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/button[2]').click()
    time.sleep(5)

    # Fecha o navegador
    navegador_chrome.quit()


def excluir_arquivo():
    # Diretório
    diretorio = 'C:\\Users\\User\\Downloads'

    # Arquivo
    arquivo = 'DIAS SEM VENDER POR CLIENTE.XLSX'

    # Caminho completo do arquivo
    caminho_arquivo = os.path.join(diretorio, arquivo)

    # Verifica se o arquivo existi e o exclui
    if os.path.exists(caminho_arquivo):
        os.remove(caminho_arquivo)



# Executa as funções
baixar_relatorio()
enviar_relatorio()
excluir_arquivo()