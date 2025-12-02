import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from urllib.parse import quote_plus
import time
import os

# CONFIGURAÇÕES
# Use nomes genéricos ou carregue de variáveis de ambiente
BASE_URL = os.getenv("TARGET_URL", "https://url-do-sistema-interno/servicos/endpoint") 
INPUT_FILE = "emails.xlsx" # Deixe o arquivo na mesma pasta do script
OUTPUT_FILE = "resultado_consulta.xlsx"

# Verifica se o arquivo existe antes de rodar
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Por favor, adicione o arquivo {INPUT_FILE} na pasta do projeto.")

# Lê os emails do Excel
df_emails = pd.read_excel(INPUT_FILE)
# Pega a primeira coluna, remove NaNs e converte em lista
emails = df_emails.iloc[:, 0].dropna().astype(str).map(str.strip).tolist()  

options = Options()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--allow-insecure-localhost')
options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(service=Service(), options=options)

dados = []

print(f"Iniciando consulta para {len(emails)} emails...")

for email in emails:
    email_encoded = quote_plus(email)
    # Monta a URL dinamicamente
    url = (f"{BASE_URL}?field_nome_sgpe_value=&field_matricula_sgpe_value="
           f"&field_email_sgpe_value={email_encoded}&items_per_page=20")
    
    driver.get(url)
    time.sleep(1) # Considerar usar WebDriverWait para produção
    
    try:
        elemento = driver.find_element(By.XPATH, '//a[@title="Ver detalhes da área"]')
        area = elemento.text
    except NoSuchElementException:
        area = "Não encontrado"
    
    dados.append({"Email": email, "Área": area})

driver.quit()

df_resultado = pd.DataFrame(dados)
df_resultado.to_excel(OUTPUT_FILE, index=False)
print(f"Consulta finalizada e arquivo salvo em {OUTPUT_FILE}")
