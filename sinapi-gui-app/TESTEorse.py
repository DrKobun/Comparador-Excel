import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import string
from tqdm import tqdm
from time import sleep
from random import uniform
from concurrent.futures import ThreadPoolExecutor, as_completed

# Período desejado
PERIODO = "2024-02-1" #Escolher período deseja
# Configurações principais
modo_teste = False  # True = teste com letra 'a' e 5 páginas #False para teste normal
TIPO_COLETA = "ambos"  # opções: "insumos", "servicos", "ambos"
MAX_WORKERS = 5  # Número de threads simultâneas

# Letras e páginas por modo
letras = ['a'] if modo_teste else list(string.ascii_lowercase)
max_paginas = 5 if modo_teste else None

def periodo_para_nome(p):
    ano, mes, versao = p.split("-")
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    nome_mes = meses[int(mes)-1]
    return f"{nome_mes}{ano}-{versao}"

def extrair_linhas(html):
    soup = BeautifulSoup(html, 'html.parser')
    linhas = []
    tabelas = soup.find_all("table")
    for tabela in tabelas:
        headers = [th.get_text(strip=True).lower() for th in tabela.find_all("td")]
        if any("descrição" in h for h in headers) and any("código" in h for h in headers):
            for row in tabela.find_all("tr")[1:]:
                cols = row.find_all("td")
                if len(cols) == 4:
                    codigo = cols[0].text.strip()
                    descricao = cols[1].text.strip()
                    unid = cols[2].text.strip()
                    valor = cols[3].text.strip().replace(",", ".")
                    linhas.append([codigo, descricao, unid, valor])
            break
    return linhas

def coletar_letra(tipo, periodo, letra, max_paginas=None):
    url_base = f"http://orse.cehop.se.gov.br/{'insumos' if tipo=='insumo' else 'servicos'}argumento.asp?tarefa=consultar"
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": url_base,
        "Origin": "http://orse.cehop.se.gov.br",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    resultado = []
    linhas_anteriores = None
    for pagina in range(1, max_paginas+1 if max_paginas else 9999):
        data = {
            "tarefa": "consultar",
            "sltFonte": "0",
            "sltPeriodo": periodo,
            "sltGrupoInsumo" if tipo == "insumo" else "sltGrupoServico": "0",
            "rdbCriterio": "1",
            "txtDescricao": letra,
            "Submit": "Consultar"
        }
        try:
            response = session.post(url_base + f"&page={pagina}", headers=headers, data=data, timeout=(10, 20))
            linhas = extrair_linhas(response.text)
            if not linhas or linhas == linhas_anteriores:
                break
            resultado.extend(linhas)
            linhas_anteriores = linhas
            sleep(uniform(0.2, 0.4))
        except Exception as e:
            print(f"⚠️ [{tipo}] Letra {letra} - Página {pagina} deu erro: {e}")
            break
    return resultado

def coletar_dados(tipo, periodo, letras, max_paginas=None):
    print(f"Coletando {tipo.upper()}S para {periodo}")
    resultados = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        tarefas = {executor.submit(coletar_letra, tipo, periodo, letra, max_paginas): letra for letra in letras}
        for fut in tqdm(as_completed(tarefas), total=len(tarefas), desc=f"{tipo.capitalize()}s"):
            try:
                resultado_letra = fut.result()
                resultados.extend(resultado_letra)
            except Exception as e:
                print(f"❌ Erro na letra {tarefas[fut]}: {e}")
    return resultados

# ▶️ Execução principal
insumos, servicos = [], []

if TIPO_COLETA in ["insumos", "ambos"]:
    insumos = coletar_dados("insumo", PERIODO, letras, max_paginas)
if TIPO_COLETA in ["servicos", "ambos"]:
    servicos = coletar_dados("servico", PERIODO, letras, max_paginas)

# DataFrames
df_insumos = pd.DataFrame(insumos, columns=["Código", "Descrição do Insumo", "Unid.", "Custo Unit."])
df_insumos.drop_duplicates(subset=["Código", "Descrição do Insumo"], inplace=True)

df_servicos = pd.DataFrame(servicos, columns=["Código", "Descrição do Serviço", "Unid.", "Custo Unit."])
df_servicos.drop_duplicates(subset=["Código", "Descrição do Serviço"], inplace=True)

# Salvar Excel
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
nome_excel = os.path.join(desktop, f"ORSE_{periodo_para_nome(PERIODO)}.xlsx")

with pd.ExcelWriter(nome_excel, engine='xlsxwriter') as writer:
    if not df_insumos.empty:
        df_insumos.to_excel(writer, sheet_name=f"{periodo_para_nome(PERIODO)} Insumos", index=False)
    if not df_servicos.empty:
        df_servicos.to_excel(writer, sheet_name=f"{periodo_para_nome(PERIODO)} Serviços", index=False)

print(f"\n Arquivo salvo em: {nome_excel}")

