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
from datetime import datetime
import sys


def executar_script():
    max_tentativas = 2
    tentativas = 0

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
            navegador = webdriver.Chrome(
                service=service, options=chrome_options)

            # Entra no gestaoclick
            navegador.get('https://gestaoclick.com.br/login/')
            time.sleep(5)

            data_atual = datetime.now().strftime("%d/%m/%Y")

            def login_gestaoclick():
                # Entra no gestão click
                navegador.find_element(
                    By.XPATH, '//*[@id="UsuarioEmail"]').send_keys('agregarnegocios@gmail.com')
                time.sleep(1)
                navegador.find_element(
                    By.XPATH, '//*[@id="UsuarioSenha"]').send_keys('khelven')
                time.sleep(1)
                navegador.find_element(
                    By.XPATH, '//*[@id="login"]/div[5]/button').click()
                time.sleep(15)
                
                # Altera a loja para "Totalmak"
                navegador.find_element(By.XPATH, '//*[@id="__BVID__42__BV_toggle_"]').click()
                navegador.find_element(By.XPATH, '//*[@id="__BVID__42__BV_toggle_menu_"]/li[2]/a').click()
                
                # Tenta logar em uma loja nova, se não, apenas mantém na que está
                try:
                    navegador.find_element(By.XPATH, '//*[@id="modal-dialog___BV_modal_footer_"]/button[2]').clic()
                except Exception as e:
                    print('Já está logado na loja')
                    
                # Aperta no botão 'Ok' para manter logado na loja
                navegador.find_element(By.XPATH, '//*[@id="modal-dialog___BV_modal_footer_"]/button')
                time.sleep(10)


            def baixar_contas(url, destinos):
                # Acessa a página do relatório
                navegador.get(url)
                time.sleep(5)

                # di = data inicio
                # df = data final
                di_contas = navegador.find_element(
                    By.XPATH, '//*[@id="data_inicio_fluxo"]')
                df_contas = navegador.find_element(
                    By.XPATH, '//*[@id="data_fim_fluxo"]')
                exibir_relatorio_detalhado = navegador.find_element(
                    By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[12]/div[1]/label')

                # Dicionário de xpaths dos centro de custos
                filtrar_centro_de_custo = [
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/ul/div/button'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/ul/li[1]/a'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/ul/li[2]/a'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/ul/li[3]/a'}
                ]

                # Filtra o período de início
                di_contas.click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                di_contas.send_keys('01/08/2023')

                # Filtra o período final
                df_contas.click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                df_contas.send_keys(data_atual)

                # Marca a opção 'Exibir relatório detalhado'
                exibir_relatorio_detalhado.click()
                time.sleep(1)

                # Repetição do processo completo de contas a pagar e a receber por centros de custo
                for i, loja in enumerate(filtrar_centro_de_custo):
                    # Abre o centro de custo
                    botao_centro_custo = navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/button')
                    botao_centro_custo.click()
                    time.sleep(1)

                    # Aperta em 'todos'
                    navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[11]/div/div/ul/div/button').click()
                    time.sleep(1)
                    botao_centro_custo.click()
                    time.sleep(1)

                    # Executa o XPath normalmente
                    navegador.find_element(
                        By.XPATH, loja['xpath_centro_custo']).click()
                    time.sleep(1)

                    # Gera o relatório
                    navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()
                    time.sleep(10)

                    # Muda o foco para a nova aba
                    navegador.switch_to.window(navegador.window_handles[-1])

                    # Exporta o relatório
                    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                    time.sleep(2)

                    # Define o nome do arquivo baixado com base na URL
                    if "contas_receber" in url:
                        downloaded_file_name = "relatorio_contas_receber.xlsx"
                    else:
                        downloaded_file_name = "relatorio_contas_pagar.xlsx"

                    # Caminho do arquivo baixado
                    downloaded_file = os.path.join(
                        download_dir, downloaded_file_name)

                    # Caminho do arquivo de destino
                    destination_file = destinos[i]['destino_conta']

                    # Verifica se o arquivo foi baixado e move para o diretório de destino
                    for _ in range(10):  # Tenta por até 10 segundos
                        if os.path.exists(downloaded_file):
                            shutil.move(downloaded_file, destination_file)
                            break
                        time.sleep(1)

                    # Fecha a página gerada para baixar o relatório
                    navegador.close()
                    time.sleep(2)

                    # Volta para a aba principal
                    navegador.switch_to.window(navegador.window_handles[0])
                    time.sleep(2)

            # Caminhos de destino para contas a pagar
            destinos_contas_pagar = [
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_contas_pagar.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_PAGAR_CURSOS.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_PAGAR_VENDAS-VAREJO.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_PAGAR_ASSISTENCIA-TECNICA.xlsx'}
            ]

            # Caminhos de destino para contas a receber
            destinos_contas_receber = [
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_contas_receber.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_RECEBER_CURSOS.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_RECEBER_VENDAS-VAREJO.xlsx'},
                {'destino_conta': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\FLUXO DE CAIXA\\CONTAS_RECEBER_ASSISTENCIA-TECNICA.xlsx'}
            ]

            def baixar_contas_receber_pagar():
                # Baixa os relatórios de contas a pagar
                baixar_contas(
                    'https://gestaoclick.com/relatorios_financeiros/relatorio_contas_pagar', destinos_contas_pagar)

                # Baixa os relatórios de contas a receber
                baixar_contas(
                    'https://gestaoclick.com/relatorios_financeiros/relatorio_contas_receber', destinos_contas_receber)

            def vendas(destinos):
                # Acessa a página de vendas
                navegador.get(
                    'https://gestaoclick.com/relatorios_vendas/relatorio_vendas')
                time.sleep(5)

                # di = data início
                navegador.find_element(
                    By.XPATH, '//*[@id="data_inicio"]').click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                navegador.find_element(
                    By.XPATH, '//*[@id="data_inicio"]').send_keys('01/08/2023')

                # df = data final
                navegador.find_element(By.XPATH, '//*[@id="data_fim"]').click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                navegador.find_element(
                    By.XPATH, '//*[@id="data_fim"]').send_keys(data_atual)
                time.sleep(1)

                # Filtra a situação para 'concretizada, concretizada'
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[10]/div/div/button').click()
                time.sleep(0.5)
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[10]/div/div/ul/li[9]/a').click()
                time.sleep(1)
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[10]/div/div/ul/li[4]/a').click()
                time.sleep(1)

                # Marca a opção 'Exibir relatório detalhado'
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[17]/div[1]/label').click()
                time.sleep(1)

                # Marca a opção "Exibir canal de venda"
                navegador.find_element(
                    By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[17]/div[3]/label').click()
                time.sleep(1)

                # Lista de xpaths dos centros de custo
                filtrar_centro_de_custo = [
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/ul/div/button'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/ul/li[1]/a'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/ul/li[2]/a'},
                    {'xpath_centro_custo':
                        '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/ul/li[3]/a'}
                ]

                # Repetição do processo completo por centros de custo
                for i, loja in enumerate(filtrar_centro_de_custo):

                    # Abre o centro de custo
                    botao_centro_custo = navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/button')
                    botao_centro_custo.click()
                    time.sleep(1)
                    # Aperta em 'todos'
                    navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[13]/div/div/ul/div/button').click()
                    time.sleep(1)
                    botao_centro_custo.click()
                    time.sleep(1)
                    navegador.find_element(
                        By.XPATH, loja['xpath_centro_custo']).click()
                    time.sleep(1)

                    # Aperta em um ponto aleatório para retirar o filtro
                    navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]').click()
                    time.sleep(1)

                    # Gera o relatório
                    navegador.find_element(
                        By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()

                    # Muda o foco para a nova aba
                    navegador.switch_to.window(navegador.window_handles[-1])

                    # Exporta o relatório
                    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                    time.sleep(2)

                    # Caminho do arquivo baixado
                    downloaded_file = os.path.join(
                        download_dir, 'relatorio_vendas.xlsx')

                    # Caminho do arquivo de destino
                    destination_file = destinos[i]['destino_venda']

                    # Verifica se o arquivo foi baixado e move para o arquivo de destino
                    for _ in range(10):  # Tenta por até 10 segundos
                        if os.path.exists(downloaded_file):
                            shutil.move(downloaded_file, destination_file)
                            break
                        time.sleep(1)

                    # Fecha a página gerada para baixar o relatório
                    navegador.close()
                    time.sleep(2)

                    # Volta para a aba principal
                    navegador.switch_to.window(navegador.window_handles[0])
                    time.sleep(2)

            # Caminhos de destino para vendas
            destinos_vendas = [
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas CURSO.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas VENDA VAREJO.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas assistencia técnica.xlsx'}
            ]

            def baixar_vendas():
                vendas(destinos_vendas)

            primeiro_dia_mes_atual = datetime.now().replace(day=1).strftime("%d/%m/%Y")

            # Caminhos de destino para vendas
            destinos_vendas = [
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas CURSO.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas VENDA VAREJO.xlsx'},
                {'destino_venda': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_vendas assistencia técnica.xlsx'}
            ]

            lista_destinos_url = [
                {'url': 'https://gestaoclick.com/relatorios_vendas/relatorio_produtos_vendidos',
                    'destino': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\PRODUTOS VENDIDOS'},
                {'url': 'https://gestaoclick.com/relatorios_vendas/relatorio_servicos_vendidos',
                    'destino': 'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\SERVIÇOS PRESTADOS'}
            ]

            def baixar_produtos_e_servicos_vendidos():
                for produtos_servicos in lista_destinos_url:
                    # Abre a página de produtos vendidos
                    navegador.get(produtos_servicos['url'])
                    time.sleep(10)

                    # di = data início
                    # df = data final
                    di_produtos = navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[3]/div/div/div[1]/div/input')
                    df_produtos = navegador.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[3]/div/div/div[4]/div/input')

                    # Filtra o período de início
                    di_produtos.click()
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(1)
                    di_produtos.send_keys(primeiro_dia_mes_atual)

                    # Filtra o período final
                    df_produtos.click()
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(1)
                    df_produtos.send_keys(data_atual)
                    time.sleep(2)

                    if 'produtos_vendidos' in produtos_servicos['url']:

                        # Filtra a situação como 'concretizada, concretizada' ná página de produtos vendidos
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[9]/div/div/button').click()
                        time.sleep(0.5)
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[9]/div/div/ul/li[4]/a').click()
                        time.sleep(1)
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[9]/div/div/ul/li[9]/a').click()
                        time.sleep(1)
                        navegador.find_element(
                            By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div').click()
                        time.sleep(1)

                    else:

                        # Filtra a situação como acima se estiver na página de serviços
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[8]/div/div/button').click()
                        time.sleep(0.5)
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[8]/div/div/ul/li[4]/a').click()
                        time.sleep(1)
                        navegador.find_element(
                            By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[8]/div/div/ul/li[9]/a').click()
                        time.sleep(1)
                        navegador.find_element(
                            By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div').click()
                        time.sleep(1)

                    # Gera o relatório
                    navegador.find_element(
                        By.XPATH, '//*[@id="app"]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()

                    # Muda o foco para a nova aba
                    navegador.switch_to.window(navegador.window_handles[-1])

                    # Exporta o relatório
                    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                    time.sleep(2)

                    # Caminho do arquivo baixado
                    if "produtos_vendidos" in produtos_servicos['url']:
                        downloaded_file = os.path.join(
                            download_dir, "relatorio_produtos_vendidos.xlsx")
                    else:
                        downloaded_file = os.path.join(
                            download_dir, "relatorio_servicos_vendidos.xlsx")

                    # Caminho do diretório de destino
                    destination_base_dir = produtos_servicos['destino']

                    # Verifica se o arquivo baixado existe
                    if os.path.exists(downloaded_file):

                        # Obtém o mês e ano atual
                        current_date = datetime.now().strftime("%m.%Y")

                        # Renomeia o arquivo baixado com o mês e ano atual
                        if "produtos_vendidos" in produtos_servicos['url']:
                            new_file_name = f"relatorio_produtos_vendidos {
                                current_date}.xlsx"
                        else:
                            new_file_name = f"resumo_servicos_vendidos {
                                current_date}.xlsx"

                        new_file_path = os.path.join(
                            destination_base_dir, new_file_name)

                        # Verifica se o arquivo renomeado já existe no diretório de destino
                        if os.path.exists(new_file_path):
                            # Se o arquivo renomeado já existir, substitui o arquivo
                            shutil.move(downloaded_file, new_file_path)
                        else:
                            # Se não, salva como novo arquivo
                            shutil.move(downloaded_file, new_file_path)

                    # Fecha a página gerada para baixar o relatório
                    navegador.close()
                    time.sleep(2)

                    # Volta para a aba principal
                    navegador.switch_to.window(navegador.window_handles[0])
                    time.sleep(2)

            def baixar_cadastros_produtos():

                # Acessa os cadastros de produtos
                navegador.get(
                    'https://gestaoclick.com/relatorios_cadastros/relatorio_produtos')
                time.sleep(5)

                # Gera o relatório
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()
                time.sleep(15)

                # Muda o foco para a nova aba
                navegador.switch_to.window(navegador.window_handles[-1])

                # Exporta o relatório
                WebDriverWait(navegador, 40).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                time.sleep(5)

                # Caminho do arquivo baixado
                downloaded_file = os.path.join(
                    download_dir, "relatorio_produtos.xlsx")

                # Caminho do arquivo de destino
                destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\PRECIFICAÇÃO\\relatorio_produtos.xlsx"

                # Move o arquivo baixado para o diretório de destino, substituindo o arquivo existente
                if os.path.exists(downloaded_file):
                    shutil.move(downloaded_file, destination_file)

                # Fecha a página gerada para baixar o relatório
                navegador.close()
                time.sleep(2)

                # Volta para a aba principal
                navegador.switch_to.window(navegador.window_handles[0])
                time.sleep(2)

            def baixar_estoque():

                # Acessa a página de estoque
                navegador.get(
                    'https://gestaoclick.com/relatorios_estoque/relatorio_estoque_produtos')
                time.sleep(10)

                # Gera o relatório
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()
                time.sleep(15)

                # Muda o foco para a nova aba
                navegador.switch_to.window(navegador.window_handles[-1])

                # Exporta o relatório
                WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                time.sleep(5)

                # Caminho do arquivo baixado
                downloaded_file = os.path.join(
                    download_dir, "relatorio_estoque_produtos.xlsx")

                # Caminho do arquivo de destino
                destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\PRECIFICAÇÃO\\relatorio_estoque_produtos.xlsx"

                # Move o arquivo baixado para o diretório de destino, substituindo o arquivo existente
                if os.path.exists(downloaded_file):
                    shutil.move(downloaded_file, destination_file)

                # Fecha a página gerada para baixar o relatório
                navegador.close()
                time.sleep(2)

                # Volta para a aba principal
                navegador.switch_to.window(navegador.window_handles[0])
                time.sleep(2)

            def baixar_compras():
                # Acessa a página de compras
                navegador.get(
                    'https://gestaoclick.com/relatorios_estoque/relatorio_compras')
                time.sleep(5)

                # Data inicial e final
                di_compras = navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[2]/div/div/div[1]/div/input')
                df_compras = navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[1]/div[2]/div/div/div[4]/div/input')

                # Filtra a data inicial
                di_compras.click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                di_compras.send_keys('01/08/2023')

                # Filtra a data final
                df_compras.click()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                df_compras.send_keys(data_atual)

                # Gera o relatório
                navegador.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/aside[2]/div/section[2]/div/div/div/form/div[2]/button[1]').click()
                time.sleep(15)

                # Muda o foco para a nova aba
                navegador.switch_to.window(navegador.window_handles[-1])

                # Exporta o relatório
                WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="dv"]/caption/button'))).click()
                time.sleep(5)

                # Caminho do arquivo baixado
                downloaded_file = os.path.join(
                    download_dir, "relatorio_compras.xlsx")

                # Caminho do arquivo de destino
                destination_file = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\relatorio_Compras.xlsx"

                # Move o arquivo baixado para o diretório de destino, substituindo o arquivo existente
                if os.path.exists(downloaded_file):
                    shutil.move(downloaded_file, destination_file)

                # Fecha a página gerada para baixar o relatório
                navegador.close()
                time.sleep(2)

                # Volta para a aba principal
                navegador.switch_to.window(navegador.window_handles[0])
                time.sleep(2)

            # Executa o login
            login_gestaoclick()

            # Baixa o contas a pagar e a receber
            baixar_contas_receber_pagar()

            # Baixa as vendas
            baixar_vendas()

            # Baixa os relatórios de produtos vendidos
            baixar_produtos_e_servicos_vendidos()

            # Baixa o relatório de produtos cadastrados
            baixar_cadastros_produtos()

            # Baixa o relatório de estoque
            baixar_estoque()

            # Baixa o relatório de compras
            baixar_compras()

            # Fecha o navegador
            navegador.quit()

            return 0  # 0 para sucesso

        except Exception as e:
            tentativas += 1
            if tentativas < max_tentativas:
                time.sleep(5)
            else:
                return 1  # para falha

        finally:
            if navegador:
                navegador.quit()
    return 1


sys.exit(executar_script())
