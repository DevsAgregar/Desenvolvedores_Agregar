import requests
import json
import pandas as pd
import time

def obter_pedidos(data_inicio, data_fim):
    pagina = 1
    todos_pedidos = []
    max_retentativas = 5  # Número máximo de retentativas
    atraso_inicial = 1  # Atraso inicial em segundos

    while True:
        url = f"https://bling.com.br/Api/v2/pedidos/page={pagina}/json"
        params = {
            'apikey': '3349e98483248cd8e598666e51cf05dfc8aa72aebf3978acec3dd3dbc13b31b1017c3eeb',
            'historico': 'true',
            'filters': f'dataEmissao[{data_inicio} TO {data_fim}]'
        }

        for tentativa in range(max_retentativas):
            response = requests.get(url, params=params)
            
            if response.status_code == 200:
                pedidos = json.loads(response.text)
                
                # Verifica se há pedidos na resposta
                if 'retorno' in pedidos:
                    if 'pedidos' in pedidos['retorno']:
                        todos_pedidos.extend(pedidos['retorno']['pedidos'])
                        pagina += 1
                        break
                    elif 'erros' in pedidos['retorno']:
                        erro = pedidos['retorno']['erros'][0]['erro']
                        if erro['cod'] == 14:
                            print("Todas as páginas foram processadas.")
                            return todos_pedidos
                else:
                    return todos_pedidos
            
            elif response.status_code == 429:
                # Atraso exponencial antes de tentar novamente
                atraso = atraso_inicial * (2 ** tentativa)
                print(f"Erro 429: Muitas solicitações. Tentativa {tentativa+1}/{max_retentativas}. Aguardando {atraso} segundos...")
                time.sleep(atraso)
            else:
                print(f"Erro ao obter os pedidos: {response.status_code}")
                return todos_pedidos

        # Aguarda para respeitar o limite de 3 requisições por segundo
        time.sleep(0.34)  # 0.34 segundos entre requisições equivale a aproximadamente 3 requisições por segundo

    return todos_pedidos

def salvar_para_excel(dados, nome_arquivo):
    if dados:
        df = pd.json_normalize(dados, sep='_')
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados salvos em {nome_arquivo}")
    else:
        print("Nenhum dado foi recuperado para salvar.")

# Exemplo de uso
data_inicio = '01/09/2024'
data_fim = '30/11/2024'
pedidos = obter_pedidos(data_inicio, data_fim)
salvar_para_excel(pedidos, 'todos_pedidos.xlsx')