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
        # Abrimos el libro que ya está configurado en settings
        with xw.App(visible=False) as app:
            wb = xw.Book(settings.Ruta_excel)
            sheet = wb.sheets["Facturas"]
            headers = sheet.range('1:1').value
            if "Descargada" in headers:
                col_index = headers.index("Descargada") + 1
                # Fila: index de pandas + 2 (1 de cabecera y 1 por base 1 de Excel)
                sheet.cells(indice_fila + 2, col_index).value = resultado
                wb.save()
                print(f"📝 Excel actualizado: Fila {indice_fila + 2} -> {resultado}")
            else:
                print("⚠️ Error: No existe la columna 'Descargada' en el Excel.")
    except Exception as e:
        print(f"❌ No se pudo actualizar el Excel: {e}")


def iniciar_driver(ruta_descargas):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
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
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def seleccionar_contrato(driver, wait, contrato_buscado):
    try:
        time.sleep(3)
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
        return False # No encontrado

    except Exception as e:
        print(f"❌ Error buscando contrato {contrato_buscado}: {e}")
        return False

def descargar_factura(driver, wait, contrato, ruta_descargas):
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame(0)

        xpath_btn = "//*[contains(@id, 'LinkCopiaCorreo')] | //*[contains(text(), 'Copia de Factura')]"
        btn_copia = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
        driver.execute_script("arguments[0].click();", btn_copia)

        time.sleep(4)
        driver.execute_script("document.body.style.zoom='70%'")

        archivos_antes = set(os.listdir(ruta_descargas))
        xpath_pdf = "//span[contains(@class, 'pi-file-pdf')]/parent::button"
       
        # Validación: ¿Existe el botón de PDF?
        try:
            btn_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_pdf)))
            driver.execute_script("arguments[0].click();", btn_pdf)
        except:
            print(f"⚠️ No se encontró botón PDF para el contrato {contrato}")
            return False

        # Espera de descarga física del archivo
        for _ in range(25):
            time.sleep(1)
            archivos_ahora = set(os.listdir(ruta_descargas))
            nuevos = archivos_ahora - archivos_antes
            finales = [f for f in nuevos if f.lower().endswith('.pdf')]
            if finales:
                destino = os.path.join(ruta_descargas, f"Factura_{contrato}.pdf")
                if os.path.exists(destino): os.remove(destino)
                os.rename(os.path.join(ruta_descargas, finales[0]), destino)
                print(f"✨ PDF Guardado: Factura_{contrato}.pdf")
                return True

        print(f"⏰ Tiempo de espera agotado para el PDF de {contrato}")
        return False
    except Exception as e:
        print(f"❌ Error en flujo de descarga: {e}")
        return False

def ejecutar_flujo():
    pendientes = leer_excel()
    if pendientes is None or pendientes.empty:
        print("No hay registros pendientes.")
        return

    driver = iniciar_driver(settings.Ruta_descargas)
    wait = WebDriverWait(driver, settings.Tiempo_espera)
    try:
        for index, row in pendientes.iterrows():
            contrato = str(row["Contrato"]).strip()
            print(f"\n🚀 PROCESANDO: {contrato}")
            try:
                # 1. Asegurar navegación
                driver.switch_to.default_content()
                if "index.xhtml" not in driver.current_url.lower():
                    driver.get(settings.Login_url)
                    time.sleep(2)
                    if len(driver.find_elements(By.NAME, "username")) > 0:
                        login(driver, row["Usuario"], row["Contraseña"])
                        time.sleep(5)
                else:
                    driver.get("https://www.acueducto.com.co/oficinavirtual/oficinavirtual/index.xhtml")
                    time.sleep(3)

                # 2. Intentar Selección
                if seleccionar_contrato(driver, wait, contrato):
                    # 3. Intentar Descarga
                    if descargar_factura(driver, wait, contrato, settings.Ruta_descargas):
                        actualizar_excel(index, "SI")
                    else:
                        actualizar_excel(index, "NO") # Falló la descarga del PDF
                else:
                    actualizar_excel(index, "NO") # No se encontró el contrato

            except Exception as e:
                print(f"⚠️ Error crítico en fila {index}: {e}")
                actualizar_excel(index, "NO")
                driver.get("https://www.acueducto.com.co/oficinavirtual/oficinavirtual/index.xhtml")
                continue
    finally:
        driver.quit()
        print(f"\n🏁 Proceso terminado.")