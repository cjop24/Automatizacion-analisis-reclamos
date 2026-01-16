import os
import pandas as pd
import logging
import time
import requests
import urllib3
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Desactivar advertencias de certificados SSL no verificados
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración de Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # Bloqueamos imágenes para mayor velocidad en el raspado
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    return webdriver.Chrome(options=options)

def download_file(url, folder, filename):
    """Descarga archivos omitiendo la verificación SSL para evitar el error de certificado."""
    try:
        if not os.path.exists(folder):
            os.makedirs(folder)
        path = os.path.join(folder, filename)
        
        # verify=False soluciona el error SSL visto en GitHub Actions
        response = requests.get(url, stream=True, timeout=20, verify=False)
        
        if response.status_code == 200:
            with open(path, 'wb') as f:
                for chunk in response.iter_content(8192):
                    f.write(chunk)
            return True
        else:
            logging.error(f"❌ Error HTTP {response.status_code} al intentar descargar: {filename}")
    except Exception as e:
        logging.error(f"⚠️ Error técnico descargando {filename}: {e}")
    return False

def run_process():
    # --- CREDENCIALES DESDE SECRETS ---
    USER = os.getenv("PQRD_USER")
    PASS = os.getenv("PQRD_PASS")
    FILE_NAME = "Reclamos.xlsx"
    
    if not USER or not PASS:
        logging.error("❌ No se encontraron las variables de entorno PQRD_USER o PQRD_PASS")
        return

    driver = get_driver()
    wait = WebDriverWait(driver, 25)

    try:
        # 1. LOGIN
        logging.info("Iniciando sesión en la plataforma...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(USER)
        driver.find_element(By.ID, "password").send_keys(PASS)
        driver.execute_script("Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR')).click();")
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso.")

        # 2. CARGA DE EXCEL
        df = pd.read_excel(FILE_NAME, dtype=str)
        
        # Asegurar columnas destino: DG(Motivos), DH(Motivos2), DI(Seguimiento), DJ(Expediente)
        cols_needed = {'Motivos': 'DG', 'Motivos2': 'DH', 'Seguimiento': 'DI', 'Expediente': 'DJ'}
        for name in cols_needed.keys():
            if name not in df.columns:
                df[name] = ""

        # 3. IDENTIFICAR PENDIENTES (Basado en Seguimiento vacío)
        df['Seguimiento'] = df['Seguimiento'].fillna("").astype(str).str.strip()
        indices_pendientes = df.index[df['Seguimiento'] == ""].tolist()
        
        logging.info(f"Registros pendientes a procesar: {len(indices_pendientes)}")

        # 4. BUCLE DE EXTRACCIÓN
        contador = 0
        for idx in indices_pendientes[:100]: # Lote de prueba de 100
            # NURC está en la columna 6 (índice 5)
            nurc = str(df.iloc[idx, 5]).strip().split('.')[0]
            if not nurc or nurc == 'nan': continue

            logging.info(f"[{contador+1}] Procesando NURC: {nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{nurc}")
            
            # Espera para renderizado de Angular
            time.sleep(7) 

            try:
                # Extracción masiva con JS
                data = driver.execute_script("""
                    let res = { motivos: "", motivos2: "", seguimiento: "", links: [] };
                    
                    // Motivos (DG)
                    let m = document.querySelector('div.ng-star-inserted p');
                    if(m) res.motivos = m.innerText.trim();

                    // Motivos 2 (DH)
                    let m2 = Array.from(document.querySelectorAll('div')).find(d => d.innerText.includes('Motivos:'));
                    if(m2) res.motivos2 = m2.innerText.replace('Motivos:', '').trim();

                    // Seguimiento (DI)
                    let rows = document.querySelectorAll('#main_table tbody tr');
                    res.seguimiento = Array.from(rows).map(r => r.innerText.replace(/\\t/g, ' ')).join('\\n---\\n');

                    // Expediente (DJ - Links)
                    res.links = Array.from(document.querySelectorAll('a[href*="anex-download"]')).map(a => a.href);
                    
                    return res;
                """)

                # Guardar en DataFrame
                df.at[idx, 'Motivos'] = data['motivos']
                df.at[idx, 'Motivos2'] = data['motivos2']
                df.at[idx, 'Seguimiento'] = data['seguimiento']
                df.at[idx, 'Expediente'] = "\n".join(data['links'])

                # 5. DESCARGA FÍSICA DE ANEXOS
                for link in data['links']:
                    filename = link.split('/')[-1]
                    download_file(link, f"anexos/{nurc}", filename)

            except Exception as e:
                logging.error(f"Error extrayendo datos de {nurc}: {e}")

            contador += 1
            # Guardado preventivo cada 10 registros
            if contador % 10 == 0:
                df.to_excel(FILE_NAME, index=False)

        # Guardado final
        df.to_excel(FILE_NAME, index=False)
        logging.info(f"✅ Proceso terminado. Total: {contador}")

    except Exception as e:
        logging.error(f"Error crítico: {e}")
        driver.save_screenshot("ERROR_CRITICO.png")
    finally:
        driver.quit()

if __name__ == "__main__":
    run_process()
