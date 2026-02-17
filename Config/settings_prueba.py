from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# 🔐 Credenciales (por ahora fijas solo para prueba)
USUARIO = "mmesa@azimutenergia.co"
PASSWORD = "Azimutenergia*"

LOGIN_URL = "https://www.acueducto.com.co/oficinavirtual/oficinavirtual/login.xhtml"

def iniciar_driver():
    chrome_options = Options()
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    driver.maximize_window()
    return driver

def login(driver, usuario, password):
    driver.get(LOGIN_URL)

    wait = WebDriverWait(driver, 15)

    # Esperar que cargue el campo usuario
    campo_usuario = wait.until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    campo_usuario.send_keys(usuario)

    driver.find_element(By.ID, "inputpass").send_keys(password)

    # Botón login
    boton_login = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
    ) 
    boton_login.click()

    print("Login ejecutado")

def main():
    driver = iniciar_driver()

    try:
        login(driver, USUARIO, PASSWORD)
        time.sleep(10)  # Solo para ver el resultado
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
