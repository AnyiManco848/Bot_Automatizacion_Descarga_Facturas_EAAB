import pandas as pd
import xlwings as xw
from Config.settings import Ruta_excel

def leer_excel():
    """
    Lee el archivo Excel usando xlwings para permitir el acceso 
    incluso si el usuario tiene el archivo abierto.
    """
    try:
        # Intentamos conectar con el archivo que ya está abierto
        # app(visible=False) permite que no se abra otra ventana de Excel
        with xw.App(visible=False) as app:
            book = xw.Book(Ruta_excel)
            sheet = book.sheets["Facturas"]
            
            # Leemos el rango usado y lo convertimos a DataFrame
            df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            
        # Limpiar espacios en nombres de columnas
        df.columns = df.columns.astype(str).str.strip()

        columna_filtro = "Factura en plataforma"
        
        if columna_filtro not in df.columns:
            print(f"⚠️ La columna '{columna_filtro}' no existe.")
            return pd.DataFrame()

        # Filtrar registros donde la columna es 'NO'
        df_pendientes = df[
            df[columna_filtro].astype(str).str.strip().str.upper() == "NO"
        ].copy()

        print(f"📋 Registros en 'NO' encontrados: {len(df_pendientes)}")
        return df_pendientes

    except Exception as e:
        print(f"❌ Error al leer el Excel con xlwings: {e}")
        return pd.DataFrame()