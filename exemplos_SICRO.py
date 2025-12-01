import requests
from bs4 import BeautifulSoup
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import re

# --- Configurações ---
BASE_URL = "https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro"
NOME_ARQUIVO_SAIDA = "LINKS_SICRO.txt"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def get_soup(url):
    """Faz a requisição com tratamento de encoding e retorna o objeto BeautifulSoup."""
    try:
        time.sleep(0.5)
        response = requests.get(url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        
        # Correção para "SOME CHARACTERS COULD NOT BE DECODED"
        response.encoding = response.apparent_encoding
        
        return BeautifulSoup(response.content, 'html.parser')
    except Exception:
        return None

def is_valid_link(href):
    """Filtra links que não são pastas de navegação."""
    if not href: return False
    if 'view' in href or 'at_download' in href: return False
    if href.endswith(('.zip', '.pdf', '.rar', '.xls', '.xlsx')): return False
    return True

def extrair_link_final(url_mes):
    """Entra na página do mês e busca o link de download direto."""
    soup = get_soup(url_mes)
    if not soup: return None
    
    # Tenta achar link de download do Plone ou um zip direto na página
    link = soup.find('a', href=re.compile(r'at_download/file|@@download'))
    if not link:
        link = soup.find('a', href=lambda x: x and x.endswith('.zip'))
    
    return link['href'] if link else None

def processar_regiao(dados_regiao, posicao_barra):
    """Função executada por cada Worker (Thread)."""
    nome_regiao, url_regiao = dados_regiao
    resultados_regiao = []
    
    soup_regiao = get_soup(url_regiao)
    if not soup_regiao: return []

    content = soup_regiao.find('div', id='content-core')
    if not content: return []

    links_estados = [a for a in content.find_all('a', href=True) if is_valid_link(a['href'])]
    
    # Barra de progresso para a Região atual
    with tqdm(total=len(links_estados), desc=f"{nome_regiao:<12}", position=posicao_barra, leave=True) as barra:
        
        for link_estado in links_estados:
            nome_estado = link_estado.get_text(strip=True)
            url_estado = link_estado['href']
            
            # --- Processar Estado ---
            soup_estado = get_soup(url_estado)
            if soup_estado:
                content_est = soup_estado.find('div', id='content-core')
                if content_est:
                    links_meses = content_est.find_all('a', href=True)
                    
                    for link_mes in links_meses:
                        texto_mes = link_mes.get_text(strip=True)
                        href_mes = link_mes['href']
                        
                        link_final = None

                        # Caso 1: O link já é o arquivo
                        if href_mes.lower().endswith(('.zip', '.rar')):
                            link_final = href_mes
                        # Caso 2: É uma página intermediária (padrão)
                        elif 'folha' in href_mes or 'relatorio' in href_mes or href_mes.endswith('/'):
                             if href_mes != url_estado:
                                link_final = extrair_link_final(href_mes)

                        if link_final:
                            resultados_regiao.append({
                                'url_download': link_final,
                                'info': f"{nome_regiao} - {nome_estado} - {texto_mes}" # Apenas para debug se precisar
                            })
            
            barra.update(1)
            
    return resultados_regiao

def main():
    print("--- Inicializando Varredura Paralela no SICRO ---")
    
    # 1. Pegar as Regiões (Thread Principal)
    soup_main = get_soup(BASE_URL)
    if not soup_main:
        print("Erro ao acessar página principal.")
        return

    content_main = soup_main.find('div', id='content-core')
    links_regioes = [a for a in content_main.find_all('a', href=True) if is_valid_link(a['href'])]
    
    dados_regioes = []
    for link in links_regioes:
        dados_regioes.append((link.get_text(strip=True), link['href']))

    print(f"Regiões identificadas: {len(dados_regioes)}\n")
    
    lista_completa_downloads = []

    # 2. Iniciar os Workers
    with ThreadPoolExecutor(max_workers=len(dados_regioes)) as executor:
        futures = []
        for i, dados in enumerate(dados_regioes):
            future = executor.submit(processar_regiao, dados, i)
            futures.append(future)
        
        for future in as_completed(futures):
            try:
                resultado = future.result()
                lista_completa_downloads.extend(resultado)
            except Exception as e:
                # O tqdm pode atrapalhar o print de erro, usamos write
                tqdm.write(f"Erro capturado em worker: {e}")

    # Pular linhas após as barras de progresso
    print("\n" * (len(dados_regioes) + 1))
    
    # 3. Gerar Arquivo TXT
    print(f"--- Processo Finalizado ---")
    print(f"Total de links encontrados: {len(lista_completa_downloads)}")
    print(f"Gerando arquivo {NOME_ARQUIVO_SAIDA}...")

    try:
        with open(NOME_ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
            for item in lista_completa_downloads:
                f.write(f"{item['url_download']}\n")
        print("Arquivo salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")

if __name__ == "__main__":
    main()