import os
import fnmatch
import win32com.client as win32
from openpyxl import Workbook

source_directory = os.path.expanduser('~/Downloads')

def process_files(file_prefix, source_directory, target_directory, target_filename_1, target_filename_2=None):
    def find_files_by_prefix(directory, prefix):
        matches = []
        for root, _, files in os.walk(directory):
            for filename in fnmatch.filter(files, f"{prefix}*.xls"):
                matches.append(os.path.join(root, filename))
        return matches

    def convert_xls_to_xlsx(input_path, output_path):
        excel = None
        try:
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False

            workbook_xls = excel.Workbooks.Open(input_path)
            sheet_xls = workbook_xls.Sheets(1)

            used_range = sheet_xls.UsedRange
            values = sheet_xls.Range(used_range.Address).Value

            workbook_xlsx = Workbook()
            sheet_xlsx = workbook_xlsx.active

            for row in values:
                sheet_xlsx.append(list(row))

            workbook_xlsx.save(output_path)
            workbook_xls.Close(SaveChanges=False)
            print(f"Arquivo convertido e salvo com sucesso em {output_path}")

        except Exception as e:
            print(f"Erro ao converter arquivo {input_path}: {e}")
        finally:
            if excel is not None:
                try:
                    excel.Application.Quit()
                except Exception as quit_e:
                    print(f"Erro ao fechar o Excel: {quit_e}")

    # Encontrar arquivos com o prefixo
    files = find_files_by_prefix(source_directory, file_prefix)

    # Ordena arquivos pela data de modificação (mais antigos primeiro)
    files.sort(key=lambda x: os.path.getmtime(x))

    # Verifica se temos pelo menos um arquivo para mover
    if not files:
        print("Não há arquivos suficientes para mover e converter.")
        return

    # Preparação do diretório de destino
    os.makedirs(target_directory, exist_ok=True)

    # Processar o arquivo mais antigo
    oldest_file = files[0]
    target_path_1 = os.path.join(target_directory, target_filename_1)
    convert_xls_to_xlsx(oldest_file, target_path_1)

    # Processar o segundo arquivo mais antigo se fornecido
    if len(files) > 1 and target_filename_2:
        second_oldest_file = files[1]
        target_path_2 = os.path.join(target_directory, target_filename_2)
        convert_xls_to_xlsx(second_oldest_file, target_path_2)

def funcao1():
    # Lógica da função 1
    print("Executando função 1")
    # Parâmetros específicos para esta função
    process_files(
        file_prefix="caixa1",
        target_directory=r'C:\Destinos\Diretorio1_Funcao1',
        target_filename_1='Caixa Financeiro Funcao1.xlsx',
        target_filename_2='Caixa Competência Funcao1.xlsx'
    )

def funcao2():
    # Lógica da função 2
    print("Executando função 2")
    # Parâmetros específicos para esta função
    process_files(
        file_prefix="caixa2",
        target_directory=r'C:\Destinos\Diretorio2_Funcao2',
        target_filename_1='Caixa Financeiro Funcao2.xlsx'
    )

def funcao3():
    # Lógica da função 3
    print("Executando função 3")
    # Parâmetros específicos para esta função
    process_files(
        file_prefix="caixa3",
        target_directory=r'C:\Destinos\Diretorio3_Funcao3',
        target_filename_1='Caixa Financeiro Funcao3.xlsx',
        target_filename_2='Caixa Competência Funcao3.xlsx'
    )

# Chamadas das funções
funcao1()
funcao2()
funcao3()