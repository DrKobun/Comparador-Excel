import webbrowser
import time
from datetime import datetime



_janeiro_2021 = False  # privada (por convenção, com "_")
_fevereiro_2021 = False
_marco_2021 = False
_abril_2021 = False


def obter_valor_janeiro_2021():
    return _janeiro_2021

def definir_valor_janeiro_2021(valor: bool):
    global _janeiro_2021
    _janeiro_2021 = valor

def obter_valor_fevereiro_2021():
    return _fevereiro_2021

def definir_valor_fevereiro_2021(valor: bool):
    global _fevereiro_2021
    _fevereiro_2021 = valor
    
def obter_valor_marco_2021():
    return _marco_2021

def definir_valor_marco_2021(valor: bool):
    global _marco_2021
    _marco_2021 = valor

def obter_valor_abril_2021():
    return _abril_2021

def definir_valor_abril_2021(valor: bool):
    global _abril_2021
    _abril_2021 = valor




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
        base_url_multiplos_meses = f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{estado}/SINAPI_ref_Insumos_Composicoes_{estado}_{ano}_"
        base_url_2017 = f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{estado}/SINAPI_ref_Insumos_Composicoes_{estado}_"
        
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
                
                
            # com retificação 01 (2022)
            elif ano == 2022 and mes in (1, 2, 3, 4):
                # exemplo: ..._082023_NaoDesonerado.zip (usando base_url_ma)
                url = base_url_ma + f"{t}Retificacao01.zip"
                links.append(url)
            
            # com retificação 02 (2022)
            elif ano == 2022 and mes == 10:
                # exemplo: ..._082023_NaoDesonerado.zip (usando base_url_ma)
                url = base_url_ma + f"{t}Retificacao02.zip"
                links.append(url)
            
            # sem retificação
            elif ano == 2022 and mes in (5,6,7,8,9,11,12):
                url = base_url_ma + f"{t}.zip"
                links.append(url)
                
                
            # 2021 ✅
            elif ano == 2021 and mes in (7, 8, 9, 10, 11, 12):
                url = base_url_ma + f"{t}.zip"
                links.append(url)
                
            # 2021 com refiticação ✅
            elif ano == 2021 and mes in (5, 6):
                url = base_url_ma + f"{t}_Retificacao01.zip"
                links.append(url)
                
            # exclusivo 2021 (1a4) ✅
            elif ano == 2021 and mes in (1,2,3,4):
                url = base_url_multiplos_meses + f"01a04" + f"_Retificacao01.zip"
                links.append(url)
                if mes == 1:
                    definir_valor_janeiro_2021(True)
                if mes == 2:
                    definir_valor_fevereiro_2021(True)
                if mes == 3:
                    definir_valor_marco_2021(True)
                if mes == 4:
                    definir_valor_abril_2021(True)
                
            
                
                
            # original:
            # https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_012022_NaoDesoneradoRetificacao01.zip
            # código:
            # https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-TO/SINAPI_ref_Insumos_Composicoes_TO_012022_NaoDesoneradoRetificacao01.zip
            
            # 2020 9a12 ✅
            elif ano == 2020 and mes in (9, 10, 11, 12):
                print("Valor de estado: ", estado)
                url = base_url_multiplos_meses + f"09a12_Retificacao01.zip"
                links.append(url)
                break
            # 2020 5a8 ✅
            elif ano == 2020 and mes in (5, 6, 7, 8):
                url = base_url_multiplos_meses + f"05a08.zip"
                links.append(url)
                break
            #2020 1a4 ✅
            elif ano == 2020 and mes in (1, 2, 3, 4):
                url = base_url_multiplos_meses + f"01a04.zip"
                links.append(url)
                break
            
            # 2019 ✅
            elif ano == 2019 and mes in (9, 10, 11, 12):
                print("Valor de estado: ", estado)
                url = base_url_multiplos_meses + f"09a12_Retificacao02.zip"
                links.append(url)
                break
            # 2019 5a8 ✅
            elif ano == 2019 and mes in (5, 6, 7, 8):
                url = base_url_multiplos_meses + f"05a08_Retificacao.zip"
                links.append(url)
                break
            # 2019 1a4 ✅
            elif ano == 2019 and mes in (1, 2, 3, 4):
                url = base_url_multiplos_meses + f"01a04_Retificacao.zip"
                links.append(url)
                break
            
            # 2018 ✅
            elif ano == 2018 and mes in (1, 2, 3, 4, 5, 6):
                # link: https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_2018_01a06.zip
                url = base_url_multiplos_meses + "01a06.zip"
                links.append(url)
                break
            # 2018 7e8 ✅
            elif ano == 2018 and mes in (7, 8):
                # link: https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_2018_07e08_Retificacao.zip
                url = base_url_multiplos_meses + "07e08_Retificacao.zip"
                links.append(url)
                break
            
            # 2018 9a12 ✅
            elif ano == 2018 and mes in (9, 10, 11, 12):
                # link: https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_2018_09a12_Retificacao.zip
                url = base_url_multiplos_meses + "09a12_Retificacao.zip"
                links.append(url)
                break
                
            # 2017 1a6 ✅
            elif ano == 2017 and mes in (1, 2, 3, 4, 5, 6):
                # link: https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_01a062017_retific.zip
                url = base_url_2017 + "01a062017_retific.zip"
                links.append(url)
                break
            # 2017 7a12 ✅
            elif ano == 2017 and mes in (7, 8, 9, 10, 11, 12):
                # link: https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-ac/SINAPI_ref_Insumos_Composicoes_AC_07a122017.zip
                url = base_url_2017 + "07a122017.zip"
                links.append(url)
                break
                
            
            else:
                url = base_url + f"{t}.zip"
                links.append(url)

    return links

def abrir_links_no_navegador(links, intervalo_segundos=1):
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