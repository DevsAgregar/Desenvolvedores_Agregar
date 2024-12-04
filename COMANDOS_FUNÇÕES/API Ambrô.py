import requests
import pandas as pd
from datetime import datetime
import time
import os

# Defina a URL e os parâmetros da API
base_url = "https://api.tiny.com.br/api2/pedidos.pesquisa.php"
tokens = [
    {"token": "69dd27d9769cc3376af13b9ad327bb439929b67d", "nome": "Token1", "nome_planilha": "Acesso Shopee"},
    {"token": "2ad6ac44ba27cc9504982c6c26f555e4e010b5d9398a52a1103b82ff02d07971", "nome": "Token2", "nome_planilha": "Acesso Shein"},
    {"token": "b1cbdd9d331fb3fddff654944bdc4de7e3adb1cefcae9f2bec9546dde73ce1a9", "nome": "Token3", "nome_planilha": "Acesso Ambro JR"}
]
data_inicial = "01/01/2024"
data_final = "31/01/2024"

def get_pedidos(token, pagina):
    params = {
        'token': token,
        'pagina': pagina,
        'formato': 'json',
        'dataInicial': data_inicial,
        'dataFinal': data_final
    }
    response = requests.get(base_url, params=params)
    
    # Verificar se a resposta da API foi bem-sucedida
    try:
        response.raise_for_status()  # Lança um erro se a requisição falhar
    except requests.exceptions.HTTPError as e:
        print(f"Falha na requisição: {e}")
        return []
    
    # Verificar se a resposta contém os dados esperados
    data = response.json()
    if 'retorno' not in data or 'pedidos' not in data['retorno']:
        print(f"Formato inesperado de resposta: {data}")
        return []
    
    return data['retorno']['pedidos']

def coletar_pedidos_por_token(token_info):
    todos_pedidos = []
    pagina = 1
    while True:
        pedidos = get_pedidos(token_info["token"], pagina)
        if not pedidos:
            break
        todos_pedidos.extend(pedidos)
        pagina += 1
    return todos_pedidos

# Medir o tempo de execução
start_time = time.time()

# Caminho para a área de trabalho
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

# Coleta e exporta pedidos para cada token
for token_info in tokens:
    pedidos = coletar_pedidos_por_token(token_info)
    df = pd.json_normalize(pedidos)
    arquivo_excel = os.path.join(desktop_path, f"{token_info['nome_planilha']}_{datetime.now().strftime('%Y%m%d')}.xlsx")
    df.to_excel(arquivo_excel, index=False)
    print(f"Dados exportados para {arquivo_excel}")

# Calcular a duração
end_time = time.time()
duration_seconds = end_time - start_time

# Converter duração para minutos ou horas, se necessário
if duration_seconds < 60:
    duration = f"{duration_seconds:.2f} segundos"
elif duration_seconds < 3600:
    duration_minutes = duration_seconds / 60
    duration = f"{duration_minutes:.2f} minutos"
else:
    duration_hours = duration_seconds / 3600
    duration = f"{duration_hours:.2f} horas"

# Arquivo de texto para tempo de execução
arquivo_tempo = os.path.join(desktop_path, "tempo_execucao.txt")

# Escreve a duração no arquivo de texto
with open(arquivo_tempo, 'w') as f:
    f.write(f"Tempo de execução: {duration}")

print(f"Tempo de execução salvo em {arquivo_tempo}")