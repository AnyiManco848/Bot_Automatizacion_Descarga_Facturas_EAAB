import time
import os
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from Config import settings
from modules.login import login
from modules.excel_reader import leer_excel

def actualizar_excel(indice_fila, resultado):
    """Escribe 'SI' o 'NO' en la columna 'Descargada' del Excel."""
    try:
        # Conectamos con el libro abierto sin crear una nueva instancia de Excel
        with xw.App(visible=False) as app:
            wb = xw.Book(settings.Ruta_excel)
            sheet = wb.sheets["Facturas"]
            headers = sheet.range('1:1').value
            if "Descargada" in headers:
                col_index = headers.index("Descargada") + 1
                # pandas index + 2 (1 por encabezado y 1 porque Excel empieza en 1)
                sheet.cells(indice_fila + 2, col_index).value = resultado
                wb.save()
                print(f"📝 Registro en Excel: Fila {indice_fila + 2} -> {resultado}")
            else:
                print("⚠️ No se encontró la columna 'Descargada'.")
    except Exception as e:
        print(f"❌ Error al actualizar Excel: {e}")

def iniciar_driver(ruta_descargas):
    chrome_options = Options()
    
    # --- MODO INVISIBLE ANTI-DETECCIÓN ---
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # User-Agent real para evitar bloqueos
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Ocultar rastro de automatización
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    if not os.path.exists(ruta_descargas):
        os.makedirs(ruta_descargas)
        
    prefs = {
        "download.default_directory": ruta_descargas,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Inyectar script para ocultar el objeto navigator.webdriver
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def seleccionar_contrato(driver, wait, contrato_buscado):
    try:
        time.sleep(4) # Espera técnica para carga de frames
        driver.switch_to.default_content()
        if len(driver.find_elements(By.TAG_NAME, "iframe")) > 0:
            driver.switch_to.frame(0)

        contrato_buscado = str(contrato_buscado).strip()
        xpath_sel = "//*[contains(@id, 'cuentasContratoUsuario')]"
        wait.until(EC.presence_of_element_located((By.XPATH, xpath_sel)))
        select_element = driver.find_element(By.XPATH, xpath_sel)

        dropdown = Select(select_element)
        encontrado = False
        for opcion in dropdown.options:
            val = opcion.get_attribute("value").strip()
            if contrato_buscado == val or contrato_buscado in opcion.text:
                if contrato_buscado == val:
                    dropdown.select_by_value(val)
                else:
                    dropdown.select_by_visible_text(opcion.text)
                encontrado = True
                break

        if encontrado:
            print(f"✅ Contrato {contrato_buscado} seleccionado.")
            time.sleep(5) 
            return True
        return False
    except Exception as e:
        print(f"❌ Error buscando contrato {contrato_buscado}: {e}")
        return False

def descargar_factura(driver, wait, contrato, ruta_descargas):
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame(0)
        
        # Botón 'Copia de Factura'
        xpath_btn = "//*[contains(@id, 'LinkCopiaCorreo')] | //*[contains(text(), 'Copia de Factura')]"
        btn_copia = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
        driver.execute_script("arguments[0].click();", btn_copia)
        
        time.sleep(6) # Tiempo para que cargue el visualizador de PDF en segundo plano
        
        archivos_antes = set(os.listdir(ruta_descargas))
        # Botón con el icono de PDF (el que realmente genera la descarga)
        xpath_pdf = "//span[contains(@class, 'pi-file-pdf')]/parent::button"
        
        try:
            btn_pdf = wait.until(EC.presence_of_element_located((By.XPATH, xpath_pdf)))
            driver.execute_script("arguments[0].click();", btn_pdf)
            print(f"⬇️ Iniciando descarga del PDF...")
        except:
            print(f"⚠️ Botón PDF no disponible para el contrato {contrato}")
            return False

        # Bucle de espera para asegurar que el archivo baje al disco
        for _ in range(30):
            time.sleep(1)
            archivos_ahora = set(os.listdir(ruta_descargas))
            nuevos = archivos_ahora - archivos_antes
            finales = [f for f in nuevos if f.lower().endswith('.pdf')]
            if finales:
                destino = os.path.join(ruta_descargas, f"Factura_{contrato}.pdf")
                if os.path.exists(destino): os.remove(destino)
                os.rename(os.path.join(ruta_descargas, finales[0]), destino)
                print(f"✨ ÉXITO: Factura_{contrato}.pdf descargada.")
                return True
        
        print("⏰ Se agotó el tiempo de descarga.")
        return False
    except Exception as e:
        print(f"❌ Error durante la descarga: {e}")
        return False

def ejecutar_flujo():
    pendientes = leer_excel()
    if pendientes is None or pendientes.empty:
        print("No se encontraron registros con 'NO' en la columna Descargada.")
        return

    print(f"🤖 Iniciando Bot en modo invisible...")
    driver = iniciar_driver(settings.Ruta_descargas)
    wait = WebDriverWait(driver, settings.Tiempo_espera)
    total_descargados = 0

    try:
        for index, row in pendientes.iterrows():
            contrato = str(row["Contrato"]).strip()
            print(f"\n--- Procesando: {contrato} ---")

            try:
                driver.switch_to.default_content()
                # Verificar si estamos logueados
                if "index.xhtml" not in driver.current_url.lower():
                    driver.get(settings.Login_url)
                    time.sleep(2)
                    if len(driver.find_elements(By.NAME, "username")) > 0:
                        login(driver, row["Usuario"], row["Contraseña"])
                        time.sleep(5)
                else:
                    driver.get("https://www.acueducto.com.co/oficinavirtual/oficinavirtual/index.xhtml")
                    time.sleep(3)

                # Intentar flujo de selección y descarga
                if seleccionar_contrato(driver, wait, contrato):
                    if descargar_factura(driver, wait, contrato, settings.Ruta_descargas):
                        actualizar_excel(index, "SI")
                        total_descargados += 1
                    else:
                        actualizar_excel(index, "NO")
                else:
                    actualizar_excel(index, "NO")

            except Exception as e:
                print(f"⚠️ Error procesando contrato {contrato}: {e}")
                actualizar_excel(index, "NO")
                driver.get("https://www.acueducto.com.co/oficinavirtual/oficinavirtual/index.xhtml")
                continue 

    finally:
        driver.quit()
        print(f"\n🏁 PROCESO TERMINADO")
        print(f"📊 Facturas obtenidas: {total_descargados}")