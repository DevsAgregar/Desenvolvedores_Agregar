import os
import shutil
from datetime import datetime

download_dir = "C:\\Users\\User\\Downloads"
# Caminho do arquivo baixado
downloaded_file = os.path.join(download_dir, "relatorio_vendas.xlsx")

# Caminho do diretório de destino
destination_base_dir = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS"

# Verifica se o arquivo baixado existe
if os.path.exists(downloaded_file):
    
    # Obtém o mês e ano atual
    current_date = datetime.now().strftime("%m_%Y")
    current_year = datetime.now().strftime("%Y")
    
    # Define o caminho do diretório do ano atual
    year_dir = os.path.join(destination_base_dir, current_year)
    
    # Verifica se a pasta do ano atual existe, se não, cria a pasta
    if not os.path.exists(year_dir):
        os.makedirs(year_dir)

    # Renomeia o arquivo baixado com o mês e ano atual
    new_file_name = f"relatorio_vendas_{current_date}.xlsx"
    new_file_path = os.path.join(year_dir, new_file_name)
  
    # Verifica se o arquivo renomeado já existe no diretório de destino
    if os.path.exists(new_file_path):
        # Se o arquivo renomeado já existir, substitui o arquivo
        shutil.move(downloaded_file, new_file_path)
    else:
        # Se não, salva como novo arquivo
        shutil.move(downloaded_file, new_file_path)