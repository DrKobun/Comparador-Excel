import os
import webbrowser
import time
from datetime import datetime

def gerar_links_sinapi(ano: int, mes: int, tipo: str):
    estados = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]
    
    if tipo not in ["Desonerado", "NaoDesonerado", "Ambos"]:
        print("Tipo inválido! Escolha 'Desonerado', 'NaoDesonerado' ou 'Ambos'.")
        return
    
    aa_mm = f"{ano}{mes:02d}"
    links = []
    
    for estado in estados:
        # https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_202305_NaoDesonerado_Retificacao01.zip
            
        
        base_url = f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{estado}/SINAPI_ref_Insumos_Composicoes_{estado}_{aa_mm}_"
        
        
            
        tipos = ["Desonerado", "NaoDesonerado"] if tipo == "Ambos" else [tipo]
        
        for t in tipos:
            if ano == 2023 and mes == 05:
                url = base_url + f"{t}_Retificacao01.zip"
                return links
                
            url = base_url + f"{t}.zip"
            links.append(url)
    
    return links

def abrir_links_no_navegador(links, intervalo_segundos=12):
    for link in links:
        print(f"Abrindo link: {link}")
        webbrowser.open(link)  # Abre o link no navegador padrão
        time.sleep(intervalo_segundos)  # Aguarda o intervalo especificado

# Exemplo de uso
if __name__ == "__main__":
    ano = 2024
    mes = 11
    tipo = "Ambos"  # Pode ser "Desonerado", "NaoDesonerado" ou "Ambos"
    
    # Gerar os links
    links = gerar_links_sinapi(ano, mes, tipo)
    
    # Abrir os links no navegador com intervalo de 5 segundos
    abrir_links_no_navegador(links, intervalo_segundos=12)
