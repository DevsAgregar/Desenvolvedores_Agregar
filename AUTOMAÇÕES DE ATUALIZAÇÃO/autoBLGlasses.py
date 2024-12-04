import os
import shutil
import glob
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
import pyautogui

def executar_script():
    max_tentativas = 2
    tentativas = 0

    # Diretório de destino único para os arquivos de caixa de competência
    diretorio_destino_caixa = "CAMINHO_DO_DIRETORIO_CAIXA"

    # Lista de credenciais e destinos
    credenciais_destinos_caixa_financeiro = [
        ("Agregar@BLGroup", "Agregar1234$", 
         ['G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA FINANCEIRO\\CAIXA FINANCEIRO BL GLASSES']
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
                navegador.get('https://www.bling.com.br/login')
                time.sleep(5)
                # Faz o login no bling
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/div[1]/input').send_keys(usuario)
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/div[2]/div/input').send_keys(senha)
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[2]/form/div/div[1]/div/button[1]').click()
                time.sleep(15)

            def baixar_caixa_financeiro(destinos, is_first_iteration):
                navegador.get('https://www.bling.com.br/caixa.php')
                time.sleep(15)
                
                # Aperta em "limpar"                     
                navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[2]/span[4]/a').click()
                time.sleep(10)
                # Abre os inputs de data             
                navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/button').click()
                time.sleep(1)
                navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/ul/li[7]').click()
                time.sleep(1)
                                                                       
                data_inicial_input = navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/input')
                data_final_input = navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[2]/div[2]/input')

                data_inicial_input.clear()
                data_inicial_input.send_keys(primeiro_dia_mes_atual)
                time.sleep(1)

                data_final_input.clear()
                data_final_input.send_keys(data_atual.strftime('%d/%m/%Y'))
                time.sleep(1)
                # Aplica os filtros de data
                navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/button').click()
                time.sleep(20)
                # Exporta o relatório
                navegador.find_element(By.XPATH, '/html/body/div[7]/div[5]/div[3]/div[1]/div[1]/div/div[3]/button[2]').click()
                time.sleep(30)

                def mover_arquivos_caixa_financeiro(download_dir, destinos, is_first_iteration):
                    lista_de_csv = sorted(
                        glob.glob(os.path.join(download_dir, '*.csv')),
                        key=os.path.getmtime
                    )

                    if lista_de_csv:
                        downloaded_file = lista_de_csv[-1]

                        if is_first_iteration:
                            arquivo_esperado = f'CAIXA FINANCEIRO BL GLASSES {data_atual.strftime("%m.%Y")}.csv'
                            destino_arquivo_esperado = os.path.join(destinos[0], arquivo_esperado)

                            if os.path.exists(destino_arquivo_esperado):
                                os.remove(destino_arquivo_esperado)
                                print(f"Arquivo existente '{arquivo_esperado}' removido para substituição.")

                            shutil.move(downloaded_file, destino_arquivo_esperado)
                            print(f"Arquivo '{downloaded_file}' movido para {destino_arquivo_esperado}")
                        else:
                            shutil.move(downloaded_file, destinos[0])
                            print(f"Arquivo '{downloaded_file}' movido para {destinos[0]}")

                mover_arquivos_caixa_financeiro(download_dir, destinos, is_first_iteration)


            # Diretório de destino único para os arquivos de caixa de competência
            diretorio_destino_caixa_competencia = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CAIXA COMPETENCIA"

            def baixar_caixa_competencia(periodo_inicio, periodo_fim):
                # Abre o Relatório de Contas Pagar Competência
                navegador.get('https://www.bling.com.br/gerenciador.relatorio.php#view/971647')
                time.sleep(10)
                # Seleciona a coluna "Competência"
                navegador.find_element(By.XPATH, '//*[@id="coluna"]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="coluna"]/option[9]').click()
                time.sleep(2)
                # Filtra o Período
                navegador.find_element(By.XPATH, '//*[@id="valor"]').send_keys(periodo_inicio.strftime('%d/%m/%Y'))
                navegador.find_element(By.XPATH, '//*[@id="valor2"]').send_keys(periodo_fim.strftime('%d/%m/%Y'))
                time.sleep(5)
                # Aplica o filtro
                navegador.find_element(By.XPATH, '//*[@id="add_filtro"]').click()
                time.sleep(10)
                # Exporta o relatório
                navegador.find_element(By.XPATH, '//*[@id="exportRelatorioLnk"]').click()
                time.sleep(15)
                navegador.find_element(By.CSS_SELECTOR, 'body > div.ui-dialog.ui-widget.ui-widget-content.ui-front.ui-dialog-buttons.ui-draggable.ui-dialog-newest > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button.Button.Button--primary.ui-button.ui-corner-all.ui-widget.button-default').click()
                time.sleep(15)
                # Reseta a página para repetir o processo
                navegador.refresh()
                time.sleep(10)

            def mover_arquivos(nome_arquivo):
                # Caminho completo do arquivo no destino
                caminho_arquivo_destino = os.path.join(diretorio_destino_caixa_competencia, nome_arquivo)

                # Ordenar arquivos por data de modificação
                arquivos = sorted([os.path.join(download_dir, f) for f in os.listdir(download_dir) if f.endswith('.csv')],
                                    key=os.path.getmtime, reverse=True)
                
                if arquivos:
                    arquivo_mais_recente = arquivos[0]

                    # Verificar se o arquivo já existe no destino
                    if os.path.exists(caminho_arquivo_destino):
                        os.remove(caminho_arquivo_destino)
                        print(f"Arquivo existente '{nome_arquivo}' removido para substituição.")

                    # Mover o arquivo para o destino
                    shutil.move(arquivo_mais_recente, caminho_arquivo_destino)
                    print(f"Arquivo '{arquivo_mais_recente}' movido para {caminho_arquivo_destino}")
                else:
                    print("Nenhum arquivo CSV encontrado para mover.")

            def executar_downloads_e_movimentacao():
                data_atual = datetime.now()

                # Definir períodos para o mês anterior e o mês atual
                primeiro_dia_mes_anterior = (data_atual.replace(day=1) - timedelta(days=1)).replace(day=1)
                ultimo_dia_mes_anterior = data_atual.replace(day=1) - timedelta(days=1)
                primeiro_dia_mes_atual = data_atual.replace(day=1)
                ultimo_dia_mes_atual = data_atual

                # Nome dos arquivos
                nome_arquivo_mes_anterior = f"CAIXA COMPETENCIA BLGROUP {primeiro_dia_mes_anterior.strftime('%m.%Y')}.csv"
                nome_arquivo_mes_atual = f"CAIXA COMPETENCIA BLGROUP {primeiro_dia_mes_atual.strftime('%m.%Y')}.csv"

                # Baixar relatórios para o mês anterior
                baixar_caixa_competencia(primeiro_dia_mes_anterior, ultimo_dia_mes_anterior)

                # Mover arquivo do mês anterior
                mover_arquivos(nome_arquivo_mes_anterior)

                # Baixar relatórios para o mês atual
                baixar_caixa_competencia(primeiro_dia_mes_atual, ultimo_dia_mes_atual)

                # Mover arquivo do mês atual
                mover_arquivos(nome_arquivo_mes_atual)

            def baixar_contas_pagar():
                # Acessa o relatório personalizado de contas a pagar
                navegador.get('https://www.bling.com.br/gerenciador.relatorio.php#view/972306')
                time.sleep(10)
                # Seleciona a coluna "Competência"
                navegador.find_element(By.XPATH, '//*[@id="coluna"]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="coluna"]/option[2]').click()
                time.sleep(2)
                # Filtra a data
                navegador.find_element(By.XPATH, '//*[@id="valor"]').send_keys(data_atual.replace(day=1).strftime('%d/%m/%Y'))
                navegador.find_element(By.XPATH, '//*[@id="valor2"]').send_keys(data_atual.strftime('%d/%m/%Y'))
                time.sleep(5)
                # Aplica a data
                navegador.find_element(By.XPATH, '//*[@id="add_filtro"]').click()
                time.sleep(10)
                # Exporta o relatório
                navegador.find_element(By.XPATH, '//*[@id="exportRelatorioLnk"]').click()
                time.sleep(5)
                navegador.find_element(By.CSS_SELECTOR, 'body > div.ui-dialog.ui-widget.ui-widget-content.ui-front.ui-dialog-buttons.ui-draggable.ui-dialog-newest > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button.Button.Button--primary.ui-button.ui-corner-all.ui-widget.button-default').click()
                time.sleep(15)

                def mover_arquivos_contas_pagar():
                    downloaded_file = os.path.join(download_dir, f"relatorio_{data_atual.strftime('%d_%m_%Y')}.csv")
                    destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\CONTAS A PAGAR"
                    new_file_name = f"CONTAS A PAGAR {data_atual.strftime('%m.%Y')}.csv"
                    new_file_path = os.path.join(destination_file, new_file_name)

                    if os.path.exists(downloaded_file):
                        shutil.move(downloaded_file, new_file_path)
                    else:
                        print(f"Arquivo '{downloaded_file}' não encontrado para mover.")

                mover_arquivos_contas_pagar()
                
            def baixar_vendas_produtos():
                # Acessa a aba "pedidos de venda"
                navegador.get('https://www.bling.com.br/exportacao.vendas.php')
                time.sleep(8)
                
                # Seleciona modelo "Padrão"
                navegador.find_element(By.XPATH, '//*[@id="spreadsheetModel"]').click()
                navegador.find_element(By.XPATH, '//*[@id="spreadsheetModel"]/option[2]').click()
                time.sleep(2)
                
                # Seleciona "todas as lojas"
                navegador.find_element(By.XPATH, '//*[@id="lojasVinculadas"]').click()
                navegador.find_element(By.XPATH, '//*[@id="todas"]').click()
                time.sleep(2)
                
                # Aperta em "Gerar planilha"
                navegador.find_element(By.XPATH, '//*[@id="btn_download"]').click()
                time.sleep(2)
                
                xpaths = [
                    '//*[@id="lista"]/tr/td/a',
                    '//*[@id="lista"]/tr[1]/td/a',
                    '//*[@id="lista"]/tr[2]/td/a'
                ]

                # Para cada XPath na lista
                for xpath in xpaths:
                    try:
                        
                        elemento_clicavel = WebDriverWait(navegador, 10).until(
                            EC.element_to_be_clickable((By.XPATH, xpath))
                        )

                        elemento_clicavel.click()
                        time.sleep(5)
        
                        break
                    except:
                        continue
                    
                def mover_arquivos_produtos(download_dir, diretorio_destino_produtos):
                    for arquivo in os.listdir(download_dir):
                        
                        if arquivo.lower().endswith('.csv'):
                            
                            caminho_origem = os.path.join(download_dir, arquivo)
                            caminho_destino = os.path.join(diretorio_destino_produtos, arquivo)
                            
                            try:
                                shutil.move(caminho_origem, caminho_destino)
                            except Exception as e:
                                print('Erro')
                                
                diretorio_destino_produtos = 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\BL GLASSES LTDA\\3. Finanças\\3 - Relatórios Financeiros\\03. BANCO DE DADOS\\PRODUTOS\\01 - PRODUTOS GERAL'
                
                mover_arquivos_produtos(download_dir, diretorio_destino_produtos)
                
            # Executa o processo para cada conjunto de credenciais e destinos
            for indice_iteracao, (usuario, senha, destinos) in enumerate(credenciais_destinos_caixa_financeiro):
                login_bling(usuario, senha)
                is_first_iteration = (indice_iteracao == 0)
                baixar_caixa_financeiro(destinos, is_first_iteration)

                if is_first_iteration:
                    baixar_vendas_produtos()
                    executar_downloads_e_movimentacao()
                    baixar_contas_pagar()

            break

        except Exception as e:
            print(f"Erro: {e}")
            tentativas += 1
            if tentativas < max_tentativas:
                time.sleep(5)

        finally:
            if navegador:
                navegador.quit()

executar_script()