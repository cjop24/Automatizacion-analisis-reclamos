import os
import pandas as pd
import logging
import time
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    return webdriver.Chrome(options=options)

def download_file(url, folder, filename):
    try:
        if not os.path.exists(folder): os.makedirs(folder)
        path = os.path.join(folder, filename)
        # Importante: Aquí podrías necesitar pasar cookies si la descarga requiere login
        response = requests.get(url, stream=True, timeout=10)
        if response.status_code == 200:
            with open(path, 'wb') as f:
                for chunk in response.iter_content(8192): f.write(chunk)
            return True
    except Exception as e:
        logging.error(f"Error descargando {filename}: {e}")
    return False

def run_process():
    USER = os.getenv("PQRD_USER")
    PASS = os.getenv("PQRD_PASS")
    FILE_NAME = "Reclamos.xlsx"
    
    driver = get_driver()
    wait = WebDriverWait(driver, 20)

    try:
        # LOGIN
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(USER)
        driver.find_element(By.ID, "password").send_keys(PASS)
        driver.execute_script("Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR')).click();")
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso.")

        # CARGA EXCEL
        df = pd.read_excel(FILE_NAME, dtype=str)
        cols = {'Motivos': 'DG', 'Motivos2': 'DH', 'Seguimiento': 'DI', 'Expediente': 'DJ'}
        for name in cols.keys():
            if name not in df.columns: df[name] = ""

        # PROCESO POR FILA
        pendientes = df.index[df['Seguimiento'] == ""].tolist()
        
        for idx in pendientes[:100]: # Lote pequeño para probar
            nurc = str(df.iloc[idx, 5]).strip().split('.')[0]
            logging.info(f"Procesando NURC: {nurc}")
            
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{nurc}")
            time.sleep(6) # Espera carga Angular
            
            # Extracción vía JS (Más robusto)
            data = driver.execute_script("""
                return {
                    motivos: document.querySelector('div.ng-star-inserted p')?.innerText || "",
                    motivos2: document.querySelector('div[class*="ng-star-inserted"] ul')?.innerText || "",
                    seguimiento: Array.from(document.querySelectorAll('#main_table tbody tr')).map(tr => tr.innerText).join('\\n'),
                    links: Array.from(document.querySelectorAll('a[href*="anex-download"]')).map(a => a.href)
                };
            """)
            
            # Guardar en DataFrame
            df.at[idx, 'Motivos'] = data['motivos']
            df.at[idx, 'Motivos2'] = data['motivos2']
            df.at[idx, 'Seguimiento'] = data['seguimiento']
            df.at[idx, 'Expediente'] = "\n".join(data['links'])
            
            # DESCARGA DE ARCHIVOS
            for link in data['links']:
                fname = link.split('/')[-1]
                download_file(link, f"anexos/{nurc}", fname)

            if idx % 10 == 0: df.to_excel(FILE_NAME, index=False)

        df.to_excel(FILE_NAME, index=False)
        logging.info("✅ Fin del proceso.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_process()
