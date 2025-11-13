from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time


# === CONFIGURAÇÕES DO SELENIUM ===
options = Options()
# options.add_argument("--headless")  # roda sem abrir janela
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# Caminho do driver (ajuste se necessário)
service = Service(r"C:\\Users\\walyson.ferreira\\Desktop\\chromedriver.exe")  # ou use o caminho absoluto

# ""

# Inicializa o navegador
driver = webdriver.Chrome(service=service, options=options)

# === LINK A SER ACESSADO ===
url = "https://www.caixa.gov.br/site/paginas/downloads.aspx"  # <-- coloque o link desejado
driver.get(url)

time.sleep(3)  # aguarda o carregamento (ajuste conforme a página)

# === LISTA PARA GUARDAR NOMES ===
sinapi_downloads = []

# === PERCORRE TODAS AS TAGS DA PÁGINA ===
all_elements = driver.find_elements(By.XPATH, "//*")

for element in all_elements:
    try:
        download_attr = element.get_attribute("download")
        if download_attr and download_attr.startswith("SINAPI"):
            sinapi_downloads.append(download_attr)
    except Exception:
        pass  # ignora elementos problemáticos

# === IMPRIME RESULTADO ===
print("Atributos 'download' encontrados que começam com 'SINAPI':")
for nome in sinapi_downloads:
    print("-", nome)

# Fecha o navegador
driver.quit()
