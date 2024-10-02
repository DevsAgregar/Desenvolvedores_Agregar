import os
import glob
import shutil
import time
from pywinauto import application
import traceback
import pandas as pd

download_dir = 'C:\\Users\\User\\Downloads'

def abrir_salvar_arquivo_excel(arquivo_xls):
    app = application.Application().start("excel.exe")
    time.sleep(1)
    
    # Encontra a janela do Excel
    janela_excel = app.window(title_re=".*Excel")
    
    # Abre o arquivo .xls
    janela_excel.wait('ready')
    janela_excel.type_keys(f'{arquivo_xls}~')  # ~ = enter
    
    time.sleep(1)
    
    # Salva o arquivo
    janela_excel.type_keys('^s')  # Ctrl + S
    janela_salvar = app.window(title='Salvar como')
    janela_salvar.wait('ready')
    janela_salvar.type_keys('{ENTER}') 
    
    app.kill()

def salvamento_de_arquivos():
    # Busca arquivos .xls no diretório
    arquivos_xls = glob.glob(os.path.join(download_dir, '*.xls'))

    # Se não houver arquivos .xls, encerra o script
    if not arquivos_xls:
        print("Nenhum arquivo .xls encontrado.")
        return

    destination_file = [
        {'destino': r'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\FINANCEIRO\\CAIXA Ambro JR.xlsx'},
        {'destino': r'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\FINANCEIRO\\caixa competencia.xlsx'},
        {'destino': r'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\MARKUP\\Banco de dados dos Produtos.xlsx'},
        {'destino': r'G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\AMBRÔ COMÉRCIO DE ROUPAS\\3. Finanças\\4 - Projeto POWERBI\\NOVO BI - LEPTK\\01. BANCO DE DADOS\\MARKUP\\relatório_de_vendas_para_análise_de_markup.xlsx'}    
    ]

    for i, arquivo_xls in enumerate(arquivos_xls):
        try:
            # Abrir e salvar o arquivo .xls com automação
            abrir_salvar_arquivo_excel(arquivo_xls)
            
            # Carrega o arquivo .xls com pandas
            df = pd.read_excel(arquivo_xls)
            
            # Define o novo nome para o arquivo .xlsx
            novo_nome = os.path.splitext(arquivo_xls)[0] + '.xlsx'  
            
            # Salva o arquivo .xlsx com pandas
            df.to_excel(novo_nome, index=False)
            
            print(f"Arquivo convertido e salvo como: {novo_nome}")
            
            # Remove o arquivo .xls original após a conversão
            os.remove(arquivo_xls)
            print(f"Arquivo XLS removido: {arquivo_xls}")
            
            # Move o arquivo convertido para o destino correto
            destino = destination_file[i % len(destination_file)]  # Usa o operador módulo para circular pelos destinos
            shutil.move(novo_nome, destino['destino'])
            print(f"Arquivo movido para: {destino['destino']}")
        
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo_xls}: {e}")
            # Aqui você pode optar por remover o arquivo .xls original se a conversão falhar
            # os.remove(arquivo_xls)
            print(traceback.format_exc())

# Chamada da função principal
salvamento_de_arquivos()
