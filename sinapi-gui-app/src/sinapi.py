import webbrowser
import time
from datetime import datetime

def gerar_links_sinapi(ano: int, mes: int, tipo: str, estados_list: list = None):
    # lista padrão se nenhum estado for informada
    estados_padrao = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]
    estados = estados_list if estados_list is not None else estados_padrao

    if tipo not in ["Desonerado", "NaoDesonerado", "Ambos"]:
        print("Tipo inválido! Escolha 'Desonerado', 'NaoDesonerado' ou 'Ambos'.")
        return []

    aa_mm = f"{ano}{mes:02d}"
    links = []

    for estado in estados:
        base_url = f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{estado}/SINAPI_ref_Insumos_Composicoes_{estado}_{aa_mm}_"
        base_url_ma = f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{estado}/SINAPI_ref_Insumos_Composicoes_{estado}_{mes:02d}{ano}_"
        
        tipos = ["Desonerado", "NaoDesonerado"] if tipo == "Ambos" else [tipo]
        for t in tipos:
            if ano == 2023 and mes == 5:
                # https://www.caixa.gov.br/Downloads/..._202305_..._Retificacao01.zip
                url = base_url + f"{t}_Retificacao01.zip"
                links.append(url)

            # verifica que ANO == 2023 E MÊS está na lista (ambas verdadeiras)
            elif ano == 2023 and mes in (1, 2, 3, 4, 6, 7, 8):
                # exemplo: ..._082023_NaoDesonerado.zip (usando base_url_ma)
                url = base_url_ma + f"{t}.zip"
                links.append(url)
            else:
                url = base_url + f"{t}.zip"
                links.append(url)

    return links

def abrir_links_no_navegador(links, intervalo_segundos=12):
    for link in links:
        print(f"Abrindo link: {link}")
        try:
            webbrowser.open(link)
        except Exception as e:
            print("Erro ao abrir link:", e)
        time.sleep(intervalo_segundos)

if __name__ == "__main__":
    ano = 2024
    mes = 11
    tipo = "Ambos"
    links = gerar_links_sinapi(ano, mes, tipo)
    abrir_links_no_navegador(links, intervalo_segundos=1)