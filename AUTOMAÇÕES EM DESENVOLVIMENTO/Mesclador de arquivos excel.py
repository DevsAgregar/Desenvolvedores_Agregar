import os
import pandas as pd
from glob import glob

def combinar_arquivos_excel(diretorio_entrada, arquivo_saida):

    todos_dataframes = []
    
    caminho_arquivos = os.path.join(diretorio_entrada, "*.xlsx")
    for arquivo in glob(caminho_arquivos):

        if "Recompras de Contrato" in os.path.basename(arquivo):
            df = pd.read_excel(arquivo)
           
            df['Nome_Arquivo'] = os.path.basename(arquivo)
            todos_dataframes.append(df)
    
    if todos_dataframes: 

        df_combinado = pd.concat(todos_dataframes, ignore_index=True)
      
        df_combinado.to_excel(arquivo_saida, index=False)


diretorio_entrada = r"G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\REFRIMAK\\3. Finanças\\4 - Power Bi\\01 - Banco de Dados\\02 - Comercial\\Recompras de Contrato - OMNI"
arquivo_saida = r"G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\REFRIMAK\\3. Finanças\\4 - Power Bi\\01 - Banco de Dados\\02 - Comercial\\Recompras de Contrato - OMNI\\Arquivo_Combinado.xlsx"

combinar_arquivos_excel(diretorio_entrada, arquivo_saida)