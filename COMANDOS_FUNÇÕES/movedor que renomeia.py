import os
import shutil
from datetime import datetime

download_dir = "C:\\Users\\User\\Downloads"
# Caminho do arquivo baixado
downloaded_file = os.path.join(download_dir, "relatorio_produtos_vendidos.xlsx")

# Caminho do diretório de destino
destination_base_dir = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\TOTAL MAK\\3. Finanças\\3 - Relatórios Financeiros\\REUNIÃO DE RESULTADOS\\PBI RADAR TOTALMAK\\01. BANCO DE DADOS\\BD"

   # Verifica se o arquivo baixado existe
if os.path.exists(downloaded_file):
    
    # Obtém o mês e ano atual
    current_date = datetime.now().strftime("%m.%Y")
    
    # Renomeia o arquivo baixado com o mês e ano atual
    new_file_name = f"relatorio_produtos_vendidos assistencia técnica {current_date}.xlsx"
    new_file_path = os.path.join(destination_base_dir, new_file_name)

    # Verifica se o arquivo renomeado já existe no diretório de destino
    if os.path.exists(new_file_path):
        # Se o arquivo renomeado já existir, substitui o arquivo
        shutil.move(downloaded_file, new_file_path)
    else:
        # Se não, salva como novo arquivo
        shutil.move(downloaded_file, new_file_path)