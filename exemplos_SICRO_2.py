import requests
from bs4 import BeautifulSoup
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import re

# --- Configurações ---
BASE_URL = "https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro"
NOME_ARQUIVO_SAIDA = "LINKS_SICRO_COMPLETO.txt"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def get_soup(url):
    """Faz a requisição com tratamento de encoding e retries."""
    for tentativa in range(3): # Tenta 3 vezes se falhar
        try:
            # Sleep aleatório pequeno para não bloquear
            # time.sleep(0.3) 
            response = requests.get(url, headers=HEADERS, timeout=40)
            response.raise_for_status()
            response.encoding = response.apparent_encoding
            return BeautifulSoup(response.content, 'html.parser')
        except Exception:
            print('teste')
            # time.sleep(1 * (tentativa + 1))
    return None

def is_valid_link(href, current_url):
    """
    Filtra links inúteis, javascript, âncoras e links de retorno.
    """
    if not href: return False
    if href.startswith(('javascript:', '#', 'mailto:')): return False
    if href == current_url: return False
    # Filtra links de sistema do Plone que não são conteúdo
    invalid_terms = ['sendto_form', 'folder_factories', 'context_state', 'login', 'createObject']
    if any(term in href for term in invalid_terms): return False
    return True

def extrair_link_download_da_pagina(soup):
    """
    Dada uma página (soup), tenta encontrar o botão de download real.
    """
    # 1. Busca por padrão Plone de download
    link = soup.find('a', href=re.compile(r'at_download/file|@@download'))
    # 2. Busca por qualquer zip na página
    if not link:
        link = soup.find('a', href=lambda x: x and x.lower().endswith(('.zip', '.rar', '.7z')))
    
    return link['href'] if link else None

def explorar_recursivamente(url_atual, nome_contexto, profundidade_max=2, nivel_atual=0):
    """
    Função recursiva que mergulha nas pastas (ex: Estado -> Ano -> Mês).
    """
    resultados = []
    
    if nivel_atual > profundidade_max:
        return []

    soup = get_soup(url_atual)
    if not soup: return []

    content = soup.find('div', id='content-core')
    if not content: return []

    links = content.find_all('a', href=True)

    for link in links:
        href = link['href']
        texto = link.get_text(strip=True)
        
        if not is_valid_link(href, url_atual): continue

        # --- CASO 1: É DIRETAMENTE UM ARQUIVO ---
        if href.lower().endswith(('.zip', '.rar', '.7z')):
            resultados.append({
                'url_download': href,
                'info': f"{nome_contexto} > {texto}"
            })
            continue

        # --- CASO 2: É UMA PÁGINA (Pode ser Mês ou Pasta de Ano) ---
        # Evita links que saem da área do site
        if BASE_URL not in href and 'gov.br/dnit' not in href: continue
        
        # Estratégia: Entramos na página para ver o que tem nela
        # Mas para não demorar uma eternidade, aplicamos heurística:
        # Se for página de "view" (comum para mês), buscamos download.
        # Se parecer um Ano (4 dígitos), mergulhamos (recursão).
        
        e_ano = texto.isdigit() and len(texto) == 4
        e_pagina_folha = 'folha' in href or 'relatorio' in href or '/view' in href
        
        # Se for um link de navegação padrão dentro do SICRO
        if e_ano or e_pagina_folha or (href.endswith('/') and href != url_atual):
            
            # Entra na página
            soup_filho = get_soup(href)
            if not soup_filho: continue

            # Tenta achar download direto lá dentro
            link_dl = extrair_link_download_da_pagina(soup_filho)
            
            if link_dl:
                # Achou arquivo!
                resultados.append({
                    'url_download': link_dl,
                    'info': f"{nome_contexto} > {texto}"
                })
            elif e_ano or (not link_dl and nivel_atual < profundidade_max):
                # Se não tem download e parece uma pasta (ex: Ano), chama a recursão
                # Isso garante que se houver "2020 -> Janeiro -> Arquivo", ele pegue.
                novos_resultados = explorar_recursivamente(href, f"{nome_contexto} > {texto}", profundidade_max, nivel_atual + 1)
                resultados.extend(novos_resultados)

    return resultados

def processar_regiao(dados_regiao, posicao_barra):
    """Worker por região."""
    nome_regiao, url_regiao = dados_regiao
    resultados_regiao = []
    
    soup_regiao = get_soup(url_regiao)
    if not soup_regiao: return []

    content = soup_regiao.find('div', id='content-core')
    if not content: return []

    # Pega lista de estados
    links_estados = [a for a in content.find_all('a', href=True) 
                     if is_valid_link(a['href'], url_regiao) and not a['href'].endswith(('.zip', '.pdf'))]
    
    with tqdm(total=len(links_estados), desc=f"{nome_regiao:<12}", position=posicao_barra, leave=True) as barra:
        for link_estado in links_estados:
            nome_estado = link_estado.get_text(strip=True)
            url_estado = link_estado['href']
            
            # CHAMA A FUNÇÃO RECURSIVA A PARTIR DO ESTADO
            # Profundidade 2 garante: Estado -> Ano -> Mês/Arquivo
            arquivos_estado = explorar_recursivamente(url_estado, f"{nome_regiao} > {nome_estado}", profundidade_max=2)
            resultados_regiao.extend(arquivos_estado)
            
            barra.update(1)
            
    return resultados_regiao

def main():
    print("--- Inicializando Varredura PROFUNDA no SICRO (Todos os anos/meses) ---")
    
    soup_main = get_soup(BASE_URL)
    if not soup_main:
        print("Erro crítico: Não foi possível acessar a URL base.")
        return

    content_main = soup_main.find('div', id='content-core')
    links_regioes = [a for a in content_main.find_all('a', href=True) 
                     if is_valid_link(a['href'], BASE_URL) and not a['href'].endswith('.zip')]
    
    dados_regioes = []
    for link in links_regioes:
        dados_regioes.append((link.get_text(strip=True), link['href']))

    print(f"Regiões identificadas: {len(dados_regioes)}")
    print("Iniciando workers... (Isso pode demorar um pouco mais devido à varredura de pastas antigas)")
    
    lista_completa_downloads = []

    # Iniciar Workers
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
                tqdm.write(f"Erro em worker: {e}")

    print("\n" * (len(dados_regioes) + 2))
    print(f"--- Processo Finalizado ---")
    print(f"Total de arquivos encontrados: {len(lista_completa_downloads)}")
    print(f"Salvando em {NOME_ARQUIVO_SAIDA}...")

    # Salva garantindo links únicos (set) para evitar duplicatas caso o site tenha links redundantes
    links_unicos = sorted(list(set(item['url_download'] for item in lista_completa_downloads)))

    try:
        with open(NOME_ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
            for link in links_unicos:
                f.write(f"{link}\n")
        print("Arquivo salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")

if __name__ == "__main__":
    main()