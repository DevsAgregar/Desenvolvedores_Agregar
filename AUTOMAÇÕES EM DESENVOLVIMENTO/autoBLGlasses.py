import os
import shutil
import glob
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from datetime import datetime
from openpyxl import load_workbook

def executar_script():
    max_tentativas = 2
    tentativas = 0

    # Lista de credenciais e destinos
    credenciais_destinos_caixa_financeiro = [
        ("Agregar@BLGroup", "Agregar1234$", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA FINANCEIRO\\CAIXA FINANCEIRO BL GLASSES.csv']
        ),

        ("consultoresagregar@b-Coltro", "Agregar1234$", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA FINANCEIRO\\CAIXA FINANCEIRO BCOLTRO.csv']
        )
    ]

    while tentativas < max_tentativas:
        navegador = None
        try:
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

            # Entra no Bling
            navegador.get('https://www.bling.com.br/login')
            time.sleep(5)

            data_atual = datetime.now()
            primeiro_dia_mes_atual = data_atual.replace(day=1).strftime('%d/%m/%Y')

            def login_bling(usuario, senha):
                # Entra no Bling
                navegador.get('https://www.bling.com.br/login')
                time.sleep(5)

                # Login no Bling
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/div[1]/input').send_keys(usuario)
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/div[2]/div/input').send_keys(senha)
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/button[1]').click()
                time.sleep(15)

            def baixar_caixa_financeiro(destinos):
                # Acessa a aba de caixas e bancos
                navegador.get('https://www.bling.com.br/caixa.php')
                time.sleep(15)

                # Clica em Limpar para limpar os filtros
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[2]/span[4]/a').click()
                time.sleep(10)

                # Abre o filtro de data e seleciona período customizado
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/button').click()
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/ul/li[7]').click()
                time.sleep(1)

                # Filtra a data 
                data_inicial_input = navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/input')
                data_final_input = navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[2]/input')

                data_inicial_input.clear()
                data_inicial_input.send_keys('01/01/2024')
                time.sleep(1)

                data_final_input.clear()
                data_final_input.send_keys(data_atual.strftime('%d/%m/%Y'))
                time.sleep(1)

                # Clica no botão de Filtrar
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/button').click()
                time.sleep(20)

                # Clica em Exportar Extrato
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[3]/button[2]').click()
                time.sleep(30)

                def mover_arquivos_caixa_financeiro(download_dir, destinos):
                    # Procura os arquivos csv baixados e ordena pela data de modificação (mais recente primeiro)
                    lista_de_csv = sorted(
                        glob.glob(os.path.join(download_dir, '*.csv')),
                        key=os.path.getmtime
                    )

                    for i, downloaded_file in enumerate(lista_de_csv):
                        if i < len(destinos):
                            if os.path.exists(downloaded_file):
                                shutil.move(downloaded_file, destinos[i])

                mover_arquivos_caixa_financeiro(download_dir, destinos)


            def atualizar_saldos(indice_iteracao):
                # Dicionário de xpaths pelo indice de iteração
                xpath_valores_por_iteracao = {
                0:{ # saldos BL
                    '//*[@id="tabela-saldo-conta"]/tr[4]/td[3]': 'B2',
                    '//*[@id="tabela-saldo-conta"]/tr[5]/td[3]': 'B3',
                    '//*[@id="tabela-saldo-conta"]/tr[6]/td[3]': 'B4',
                    '//*[@id="tabela-saldo-conta"]/tr[7]/td[3]': 'B5',
                    '//*[@id="tabela-saldo-conta"]/tr[8]/td[3]': 'B6',
                    '//*[@id="tabela-saldo-conta"]/tr[12]/td[3]': 'B7'
                },
                
                1:{ # saldos bcoltro
                    '//*[@id="tabela-saldo-conta"]/tr[3]/td[3]': 'B8'
                }
                }
               
                caminho_planilha = 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA FINANCEIRO\\SALDO DAS CONTAS.xlsx'
                wb = load_workbook(caminho_planilha)
                sheet = wb.active
                
                xpath_valor_excel = xpath_valores_por_iteracao[indice_iteracao]
                
                # itera sobre a lista de xpaths e insere os valores na planilha
                for xpath, celula in xpath_valor_excel.items():
                    try:
                        elemento = navegador.find_element(By.XPATH, xpath)
                        valor = elemento.text.strip()
                        sheet[celula] = valor
                    except Exception as e:
                        print('erro')
                        
                # Salva a planilha
                wb.save(caminho_planilha)   
                
                
            def baixar_caixa_competencia():
                # Acessa o relatório personalizado "Caixa Competência"
                navegador.get('https://www.bling.com.br/gerenciador.relatorio.php#view/971647')
                time.sleep(10)
                
                # Seleciona a coluna "Competência" para filtrar a data
                navegador.find_element(By.XPATH, '//*[@id="coluna"]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="coluna"]/option[9]').click()
                time.sleep(2)
                
                # Filtra o período
                navegador.find_element(By.XPATH, '//*[@id="valor"]').send_keys(data_atual.replace(day=1).strftime('%d/%m/%Y'))
                navegador.find_element(By.XPATH, '//*[@id="valor2"]').send_keys(data_atual.strftime('%d/%m/%Y'))
                time.sleep(5)
                
                # Aperta no botão de filtrar
                navegador.find_element(By.XPATH, '//*[@id="add_filtro"]').click()
                time.sleep(10)
                
                # Aperta na opção "Exportar" e no botão "Exportar" no pop-up que se abre
                navegador.find_element(By.XPATH, '//*[@id="exportRelatorioLnk"]').click()
                time.sleep(5)
                navegador.find_element(By.CSS_SELECTOR, 'body > div.ui-dialog.ui-widget.ui-widget-content.ui-front.ui-dialog-buttons.ui-draggable.ui-dialog-newest > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button.Button.Button--primary.ui-button.ui-corner-all.ui-widget.button-default').click()
                time.sleep(15)
                
                def mover_arquivos_caixa_competencia():
                    
                    downloaded_file = os.path.join(download_dir, f"relatorio_{data_atual.strftime('%d_%m_%Y')}.csv")                   
                    destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA COMPETENCIA"
                    new_file_name = f"CAIXA COMPETENCIA BLGROUP {data_atual.strftime('%m.%Y')}.csv"
                    new_file_path = os.path.join(destination_file, new_file_name)
                    
                    # Move os arquivos
                    if os.path.exists(downloaded_file):
                        shutil.move(downloaded_file, new_file_path)
                    else:
                        shutil.move(downloaded_file, new_file_path)
                    
                mover_arquivos_caixa_competencia()
                    
                    
            def baixar_contas_pagar():
                # Acessa o relatório personalizado "Caixa Competência"
                navegador.get('https://www.bling.com.br/gerenciador.relatorio.php#view/972306')
                time.sleep(10)
                
                # Seleciona a coluna "Competência" para filtrar a data
                navegador.find_element(By.XPATH, '//*[@id="coluna"]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="coluna"]/option[2]').click()
                time.sleep(2)
                
                # Filtra o período
                navegador.find_element(By.XPATH, '//*[@id="valor"]').send_keys(data_atual.replace(day=1).strftime('%d/%m/%Y'))
                navegador.find_element(By.XPATH, '//*[@id="valor2"]').send_keys(data_atual.strftime('%d/%m/%Y'))
                time.sleep(5)
                
                # Aperta no botão de filtrar
                navegador.find_element(By.XPATH, '//*[@id="add_filtro"]').click()
                time.sleep(10)
                
                # Aperta na opção "Exportar" e no botão "Exportar" no pop-up que se abre
                navegador.find_element(By.XPATH, '//*[@id="exportRelatorioLnk"]').click()
                time.sleep(5)
                navegador.find_element(By.CSS_SELECTOR, 'body > div.ui-dialog.ui-widget.ui-widget-content.ui-front.ui-dialog-buttons.ui-draggable.ui-dialog-newest > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button.Button.Button--primary.ui-button.ui-corner-all.ui-widget.button-default').click()
                time.sleep(15)
                
                def mover_arquivos_contas_pagar():
                    
                    downloaded_file = os.path.join(download_dir, f"relatorio_{data_atual.strftime('%d_%m_%Y')}.csv")                   
                    destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CONTAS A PAGAR"
                    new_file_name = f"CONTAS A PAGAR {data_atual.strftime('%m.%Y')}.csv"
                    new_file_path = os.path.join(destination_file, new_file_name)
                    
                    # Move os arquivos
                    if os.path.exists(downloaded_file):
                        shutil.move(downloaded_file, new_file_path)
                    else:
                        shutil.move(downloaded_file, new_file_path)
                    
                mover_arquivos_contas_pagar()
                

            funcao_executada = False
            
            # Executa o processo para cada conjunto de credenciais e destinos
            for indice_iteracao, (usuario, senha, destinos) in enumerate(credenciais_destinos_caixa_financeiro):
                login_bling(usuario, senha)
                baixar_caixa_financeiro(destinos)
                atualizar_saldos(indice_iteracao)

                if not funcao_executada:
                    baixar_caixa_competencia()
                    baixar_contas_pagar()
                    funcao_executada = True
                    
            break
            
        except Exception as e:
            tentativas += 1
            if tentativas < max_tentativas:
                time.sleep(5)
        
        finally:
            if navegador:
                navegador.quit()

executar_script()