import os
import pandas as pd
import logging
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # Bloqueamos imágenes para mayor velocidad
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    return webdriver.Chrome(options=options)

def extract_data_js(driver):
    """
    Inyecta JavaScript para extraer los 4 campos requeridos de forma atómica.
    Esto es más rápido y evita errores de sincronización de Selenium.
    """
    script = """
    let results = { motivos: "", motivos2: "", seguimiento: "", expediente: "" };

    // 1. MOTIVOS (Columna DG)
    let descElem = document.querySelector('div.ng-star-inserted p strong')?.parentElement;
    if (descElem) results.motivos = descElem.innerText.trim();

    // 2. MOTIVOS 2 (Columna DH)
    let mot2Elem = Array.from(document.querySelectorAll('div')).find(d => d.innerText.includes('Motivos:'));
    if (mot2Elem) results.motivos2 = mot2Elem.innerText.replace('Motivos:', '').trim();

    // 3. SEGUIMIENTO (Columna DI)
    let rows = document.querySelectorAll('#main_table tbody tr');
    if (rows.length > 0 && !rows[0].innerText.includes('No hay datos')) {
        results.seguimiento = Array.from(rows).map(r => {
            let cols = r.querySelectorAll('td');
            return cols.length >= 4 ? `[${cols[0].innerText}] ${cols[2].innerText}: ${cols[3].innerText}` : "";
        }).join('\\n---\\n');
    }

    // 4. EXPEDIENTE (Columna DJ - Links de descarga)
    let links = document.querySelectorAll('div[col-id="anex_nomb_archivo"] a');
    if (links.length > 0) {
        results.expediente = Array.from(links).map(a => a.href).join('\\n');
    }

    return results;
    """
    return driver.execute_script(script)

def run_process():
    # --- CONFIGURACIÓN DE SECRETS ---
    USER = os.getenv("PQRD_USER")
    PASS = os.getenv("PQRD_PASS")
    FILE_NAME = "Reclamos.xlsx"
    
    driver = get_driver()
    wait = WebDriverWait(driver, 20)

    try:
        # LOGIN
        logging.info("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(USER)
        driver.find_element(By.ID, "password").send_keys(PASS)
        driver.execute_script("Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR')).click();")
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso.")

        # CARGA DE EXCEL
        df = pd.read_excel(FILE_NAME, dtype=str)
        
        # Mapeo de columnas (Aseguramos que existan o las creamos)
        cols_map = {'DG': 'Motivos', 'DH': 'Motivos2', 'DI': 'Seguimiento', 'DJ': 'Expediente'}
        for col_idx, (letter, name) in enumerate(cols_map.items()):
            # Si el DataFrame es más pequeño que la posición de la columna, expandimos
            if name not in df.columns:
                df[name] = ""

        # PROCESAMIENTO
        # El NURC está en la columna 6 (índice 5)
        indices_pendientes = df.index[df['Seguimiento'].isna() | (df['Seguimiento'] == "")].tolist()
        
        for i, idx in enumerate(indices_pendientes[:500]): # Lote de 500
            nurc = str(df.iloc[idx, 5]).strip().split('.')[0] # Limpia el .0 si existe
            
            logging.info(f"[{i+1}/{len(indices_pendientes)}] Procesando NURC: {nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{nurc}")
            
            # Espera a que la tabla o el contenido cargue (Angular)
            time.sleep(5) 
            
            try:
                data = extract_data_js(driver)
                df.at[idx, 'Motivos'] = data['motivos']
                df.at[idx, 'Motivos2'] = data['motivos2']
                df.at[idx, 'Seguimiento'] = data['seguimiento']
                df.at[idx, 'Expediente'] = data['expediente']
            except Exception as e:
                logging.warning(f"Error extrayendo NURC {nurc}: {e}")

            # Guardado preventivo cada 20
            if i % 20 == 0:
                df.to_excel(FILE_NAME, index=False)

        df.to_excel(FILE_NAME, index=False)
        logging.info("✅ Proceso completado exitosamente.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_process()
