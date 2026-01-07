import os
import xlwings as xw

"""
Abre um arquivo Excel, converte todas as fórmulas em todas as planilhas
para seus valores literais e salva o arquivo.
:param file_path: O caminho para o arquivo Excel a ser processado.
"""
file_path = "C:\\Users\\walyson.ferreira\\Downloads\\Planilha_Orçamentária.xlsx"
if not os.path.isfile(file_path):
    print(f"Erro: Arquivo não encontrado em '{file_path}'")
    
print(f"Processando arquivo: {os.path.basename(file_path)}")

# Usar visible=False para rodar em segundo plano
with xw.App(visible=False) as app:
    wb = None
    try:
        wb = app.books.open(file_path)
        
        if wb.api.ReadOnly:
            print(f"  - AVISO: O arquivo está no modo somente leitura. As alterações não serão salvas.")
        for sheet in wb.sheets:
            print(f"  - Convertendo fórmulas na planilha '{sheet.name}'...")
            try:
                # A forma mais eficiente com xlwings é ler os valores e escrevê-los de volta.
                # Isso força a avaliação das fórmulas pelo Excel e as substitui pelos resultados.
                sheet.used_range.value = sheet.used_range.value
            except Exception as e:
                print(f"    - AVISO ao processar a planilha '{sheet.name}': {e}")
        print(f"Salvando alterações...")
        wb.save()
        print("Arquivo salvo com sucesso.")
    except Exception as e:
        print(f"ERRO: Não foi possível processar o arquivo. Motivo: {e}")
    finally:
        if wb:
            wb.close()

print("\nProcesso concluído.")