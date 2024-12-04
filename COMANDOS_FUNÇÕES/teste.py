import win32com.client as win32
from openpyxl import Workbook

def convert_xls_to_xlsx(input_path, output_path):
    excel = None
    try:
        # Inicia o Excel via COM
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # Modo headless

        # Abre o arquivo .xls
        workbook_xls = excel.Workbooks.Open(input_path)
        sheet_xls = workbook_xls.Sheets(1)

        # Determina o alcance utilizado na planilha
        used_range = sheet_xls.UsedRange
        num_rows = used_range.Rows.Count
        num_cols = used_range.Columns.Count

        # Obtém todos os valores de uma vez
        values = sheet_xls.Range(used_range.Address).Value

        # Cria um novo arquivo .xlsx com openpyxl
        workbook_xlsx = Workbook()
        sheet_xlsx = workbook_xlsx.active

        # Adiciona os dados ao novo arquivo .xlsx
        for row in values:
            sheet_xlsx.append(list(row))

        # Salva o novo arquivo .xlsx
        workbook_xlsx.save(output_path)

        workbook_xls.Close(SaveChanges=False)
        print(f"Arquivo convertido e salvo com sucesso em {output_path}")

    except Exception as e:
        print(f"Erro ao converter arquivo: {e}")
    finally:
        if excel is not None:
            try:
                excel.Application.Quit()
            except Exception as quit_e:
                print(f"Erro ao fechar o Excel: {quit_e}")

# Caminho para o arquivo xls
input_file = r'G:\Drives compartilhados\Agregar Negócios - Drive Geral\Agregar Clientes Ativos\AMBRÔ COMÉRCIO DE ROUPAS\3. Finanças\4 - Projeto POWERBI\01. BANCO DE DADOS\FINANCEIRO\caixa_2024-11-01-08-06-35.xls'

# Caminho para o novo arquivo xlsx
output_file = r'G:\Drives compartilhados\Agregar Negócios - Drive Geral\Agregar Clientes Ativos\AMBRÔ COMÉRCIO DE ROUPAS\3. Finanças\4 - Projeto POWERBI\01. BANCO DE DADOS\FINANCEIRO\Caixa Competência.xlsx'

# Chamar a função para converter
convert_xls_to_xlsx(input_file, output_file)