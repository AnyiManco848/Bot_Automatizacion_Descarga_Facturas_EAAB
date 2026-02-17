import subprocess
import sys

def instalar_dependencias():
    try:
        import selenium
        import pandas
    except ImportError:
        print("🔧 Preparando herramientas necesarias (esto solo tardará un momento)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "pandas", "webdriver-manager", "openpyxl", "xlwings"])
        print("✅ Herramientas listas.")

if __name__ == "__main__":
    # Primero nos aseguramos de que todo esté instalado
    instalar_dependencias()
    
    # Ahora sí, cargamos el flujo y lo ejecutamos
    from modules.flujo_eaab import ejecutar_flujo
    print("🚀 Iniciando el proceso de facturación...")
    ejecutar_flujo()