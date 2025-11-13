import openpyxl
import os

# Caminho do arquivo
caminho = r"C:\Users\walyson.ferreira\Desktop\SinapiDownloads\aninhar\aninhado_20251113_101941.xlsx"
contagem_abas = 0
contagem_des = 0
contagem_nds = 0

# Verifica se o arquivo existe
if not os.path.exists(caminho):
    print("Arquivo não encontrado:", caminho)
else:
    # Carrega o arquivo Excel
    workbook = openpyxl.load_workbook(caminho)

    # Lista os nomes das abas
    abas = workbook.sheetnames
    
    # abas é tipo lista
    print("Tipo de abas: ", type(abas))
    
    for aba in abas:
        if "NDS" in aba:
            contagem_nds = contagem_nds + 1
        elif "DES" in aba:
            contagem_des = contagem_des + 1
            

    for aba in abas:
        contagem_abas = contagem_abas + 1
    
    # Exibe o resultado no terminal
    print("Abas encontradas no arquivo:")
    for nome in abas:
        print(f"- {nome}")
    print("Total de abas: ", contagem_abas)
    print("Total de abas Desoneradas: ", contagem_des)
    print("Total de abas Não Desoneradas: ", contagem_nds)
