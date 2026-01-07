import openpyxl
import os

# 1. Defina o caminho do seu arquivo aqui
caminho_arquivo = 'C:\\Users\\walyson.ferreira\\Desktop\\Arquivos-SINAPI-SICRO-ORSE\\agrupado\\agrupado_20260107_085837.xlsx'
# caminho_saida = 'C:\\Users\\walyson.ferreira\\Desktop\\Arquivos-SINAPI-SICRO-ORSE\\agrupado\\teste.xlsx'

def corrigir_excel_no_proprio_arquivo(caminho):
    if not os.path.exists(caminho):
        print("Erro: Arquivo não encontrado.")
        return

    print(f"Abrindo arquivo: {caminho}")
    # Carrega o workbook (arquivo) mantendo estilos e fórmulas
    wb = openpyxl.load_workbook(caminho)

    for sheet in wb.worksheets:
        print(f"Processando aba: {sheet.title}")
        
        # Percorre todas as células da planilha
        for row in sheet.iter_rows():
            for cell in row:
                # Verifica se a célula tem um valor e se esse valor é uma string
                if cell.value is not None and isinstance(cell.value, str):
                    valor_limpo = cell.value.strip()
                    
                    # Tenta converter para número (float ou int)
                    try:
                        # Se a conversão for bem-suceda, o openpyxl remove o apóstrofo
                        if '.' in valor_limpo or ',' in valor_limpo:
                            # Trata caso de vírgula como separador decimal
                            valor_num = float(valor_limpo.replace(',', '.'))
                        else:
                            valor_num = int(valor_limpo)
                        
                        cell.value = valor_num
                        # Opcional: define o formato de número geral para garantir
                        cell.number_format = 'General'
                    except ValueError:
                        # Se não for um número (ex: nomes, endereços), mantém como está
                        continue

    # 2. Salva no mesmo caminho (sobrescrevendo o original)
    try:
        wb.save(caminho)
        print("\nSucesso! O erro de 'número como texto' foi corrigido no próprio arquivo.")
        print("Largura de colunas e nomes originais foram preservados.")
    except PermissionError:
        print("\nErro: O arquivo está aberto. Feche o Excel e tente novamente.")
    except Exception as e:
        print(f"\nOcorreu um erro ao salvar: {e}")

# Executar a função
corrigir_excel_no_proprio_arquivo(caminho_arquivo)