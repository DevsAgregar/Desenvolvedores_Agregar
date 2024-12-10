import requests
import pandas as pd
from datetime import datetime
import time
import os

data_inicial = "01/" + datetime.now().strftime("%m/%Y")
data_final = f"{datetime.now().strftime('%d/%m/%Y')}"

data_inicial_obj = datetime.strptime(data_inicial, "%d/%m/%Y")
data_formatada = data_inicial_obj.strftime("%m.%Y")

base_url = "https://api.tiny.com.br/api2/pedidos.pesquisa.php"

tokens = [
    {"token": "69dd27d9769cc3376af13b9ad327bb439929b67d", "nome": "Token1", "nome_planilha": f"Acesso Shopee {data_formatada}"},
    {"token": "b1cbdd9d331fb3fddff654944bdc4de7e3adb1cefcae9f2bec9546dde73ce1a9", "nome": "Token2", "nome_planilha": f"Acesso Ambro JR {data_formatada}"}
]

destinos = {
    "Token1": r"G:\Drives compartilhados\Agregar Negócios - Drive Geral\Agregar Clientes Ativos\AMBRÔ COMÉRCIO DE ROUPAS\3. Finanças\4 - Projeto POWERBI\01. BANCO DE DADOS\02 - COMERCIAL\2.1 - VENDAS\ACESSO SHOPEE",
    "Token2": r"G:\Drives compartilhados\Agregar Negócios - Drive Geral\Agregar Clientes Ativos\AMBRÔ COMÉRCIO DE ROUPAS\3. Finanças\4 - Projeto POWERBI\01. BANCO DE DADOS\02 - COMERCIAL\2.1 - VENDAS\ACESSO AMBRO JR"
}

def get_pedidos(token, pagina):
    params = {
        'token': token,
        'pagina': pagina,
        'formato': 'json',
        'dataInicial': data_inicial,
        'dataFinal': data_final
    }
    response = requests.get(base_url, params=params)
    
    try:
        response.raise_for_status()
    except requests.exceptions.HTTPError as e:
        print(f"Falha na requisição: {e}")
        return []
    
    data = response.json()
    
    if 'retorno' in data:
        if 'status' in data['retorno'] and data['retorno']['status'] == 'Erro':
            print(f"Erro ao consultar API: {data['retorno']}")
            return []
        elif 'pedidos' in data['retorno']:
            return data['retorno']['pedidos']
    
    print(f"Formato inesperado de resposta: {data}")
    return []

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

start_time = time.time()

for token_info in tokens:
    destino = destinos.get(token_info["nome"])
    
    if destino is None:
        print(f"Destino não encontrado para {token_info['nome']}. Pulando...")
        continue
    
    if not os.path.exists(destino):
        os.makedirs(destino)
    
    pedidos = coletar_pedidos_por_token(token_info)
    
    if not pedidos:
        print(f"Nenhum pedido encontrado para {token_info['nome']}. Pulando a exportação.")
        continue
    
    df = pd.json_normalize(pedidos)
    arquivo_excel = os.path.join(destino, f"{token_info['nome_planilha']}.xlsx")
    df.to_excel(arquivo_excel, index=False)
    print(f"Dados exportados para {arquivo_excel}")

end_time = time.time()
duration_seconds = end_time - start_time

if duration_seconds < 60:
    duration = f"{duration_seconds:.2f} segundos"
elif duration_seconds < 3600:
    duration_minutes = duration_seconds / 60
    duration = f"{duration_minutes:.2f} minutos"
else:
    duration_hours = duration_seconds / 3600
    duration = f"{duration_hours:.2f} horas"

print(f"Tempo de execução: {duration}")