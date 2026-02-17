import time
import random
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def escribir_como_humano(elemento, texto):
    elemento.clear()
    for letra in str(texto):
        elemento.send_keys(letra)
        time.sleep(random.uniform(0.05, 0.15))

def login(driver, usuario, password):
    wait = WebDriverWait(driver, 30)
    driver.get("https://www.acueducto.com.co/oficinavirtual/oficinavirtual/login.xhtml")
    
    user_field = wait.until(EC.element_to_be_clickable((By.NAME, "username")))
    escribir_como_humano(user_field, usuario)
    escribir_como_humano(driver.find_element(By.NAME, "password"), password)
    
    driver.find_element(By.XPATH, "//button[contains(.,'Ingresar')]").click()
    
    # Esperar a que la URL cambie al index
    wait.until(EC.url_contains("index.html"))
    time.sleep(5)