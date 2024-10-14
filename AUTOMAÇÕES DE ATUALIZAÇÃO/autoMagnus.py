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
    credenciais_destinos = [
        ("ksoares@MAGNUSCABELOS", "Ksoares2024.", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\MAGNUS CABELOS\\3. Finanças\\4 - Projeto BI\\01. BANCO DE DADOS\\CAIXA ATUAL\\CAIXA MAGNUS.csv']
        ),

        ("ksoares@imperadorc", "Ksoares2024.", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\MAGNUS CABELOS\\3. Finanças\\4 - Projeto BI\\01. BANCO DE DADOS\\CAIXA ATUAL\\CAIXA IMPERADOR.csv']
        ),

        ("ksoares@imperadorio", "Ksoares2024.", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\MAGNUS CABELOS\\3. Finanças\\4 - Projeto BI\\01. BANCO DE DADOS\\CAIXA ATUAL\\CAIXA IMPERADOR RIO.csv']
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

            data_atual = datetime.now().strftime('%d/%m/%Y')

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

            def baixar_caixa():
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
                data_final_input.send_keys(data_atual)
                time.sleep(1)

                # Clica no botão de Filtrar
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/button').click()
                time.sleep(20)

                # Clica em Exportar Extrato
                navegador.find_element(By.XPATH, '/html/body/div[6]/div[5]/div[3]/div[1]/div[1]/div/div[3]/button[2]').click()
                time.sleep(30)

            def atualizar_saldos(navegador, indice_iteracao):
                # Dicionário de xpaths pelo indice de iteração
                xpath_valores_por_iteracao = {
                0:{ # saldos magnus
                    '//*[@id="tabela-saldo-conta"]/tr[1]/td[3]': 'B2',
                    '//*[@id="tabela-saldo-conta"]/tr[4]/td[3]': 'B3',
                    '//*[@id="tabela-saldo-conta"]/tr[5]/td[3]': 'B4',
                    '//*[@id="tabela-saldo-conta"]/tr[10]/td[3]': 'B5',
                    '//*[@id="tabela-saldo-conta"]/tr[11]/td[3]': 'B6',
                    '//*[@id="tabela-saldo-conta"]/tr[12]/td[3]': 'B7',
                    '//*[@id="tabela-saldo-conta"]/tr[13]/td[3]': 'B8',
                    '//*[@id="tabela-saldo-conta"]/tr[14]/td[3]': 'B9',
                    '//*[@id="tabela-saldo-conta"]/tr[17]/td[3]': 'B10',
                    '//*[@id="tabela-saldo-conta"]/tr[18]/td[3]': 'B11',
                    '//*[@id="tabela-saldo-conta"]/tr[19]/td[3]': 'B12',
                },
                1:{ # saldos imperador
                    '//*[@id="tabela-saldo-conta"]/tr[1]/td[3]': 'C2',
                    '//*[@id="tabela-saldo-conta"]/tr[5]/td[3]': 'C5',
                    '//*[@id="tabela-saldo-conta"]/tr[6]/td[3]': 'C13',
                    '//*[@id="tabela-saldo-conta"]/tr[7]/td[3]': 'C14',
                    '//*[@id="tabela-saldo-conta"]/tr[8]/td[3]': 'C16',
                    '//*[@id="tabela-saldo-conta"]/tr[9]/td[3]': 'C10',
                    '//*[@id="tabela-saldo-conta"]/tr[10]/td[3]': 'C15'
                },

                2:{ # saldos rio
                    '//*[@id="tabela-saldo-conta"]/tr[1]/td[3]': 'D2',
                    '//*[@id="tabela-saldo-conta"]/tr[4]/td[3]': 'D5',
                    '//*[@id="tabela-saldo-conta"]/tr[5]/td[3]': 'D10',
                    '//*[@id="tabela-saldo-conta"]/tr[6]/td[3]': 'D15'
                }
                }

                # Abre a planilha de saldos
                caminho_planilha = 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\MAGNUS CABELOS\\3. Finanças\\4 - Projeto BI\\01. BANCO DE DADOS\\SALDO DAS CONTAS.xlsx'
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


            def mover_arquivos(download_dir, destinos):
                # Procura os arquivos CSV baixados e ordena pela data de modificação (mais recente primeiro)
                lista_de_csv = sorted(
                    glob.glob(os.path.join(download_dir, '*.csv')),
                    key=os.path.getmtime
                )

                if not lista_de_csv:
                    return

                # Move os arquivos baixados para os destinos
                for i, downloaded_file in enumerate(lista_de_csv):
                    if i < len(destinos):

                        if os.path.exists(downloaded_file):
                            shutil.move(downloaded_file, destinos[i])

            # Executa o processo para cada conjunto de credenciais e destinos
            for indice_iteracao, (usuario, senha, destinos) in enumerate(credenciais_destinos):
                login_bling(usuario, senha)
                baixar_caixa()
                atualizar_saldos(navegador, indice_iteracao)
                mover_arquivos(download_dir, destinos)

            break

        except Exception as e:
            tentativas += 1
            if tentativas < max_tentativas:
                time.sleep(5)
        
        finally:
            if navegador:
                navegador.quit()

executar_script()