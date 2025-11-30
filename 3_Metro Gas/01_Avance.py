import os
import time
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================================
# GOOGLE DRIVE (TOKEN PERSISTENTE)
# ================================
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

CREDENCIALES = "key_drive.json"
ID_CARPETA_DRIVE = "1BCb9JXw17Eix4Ux2BIzN1ok50lgFLb_H"


def login_drive():
    gauth = GoogleAuth()
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = CREDENCIALES

    gauth.LoadCredentialsFile(CREDENCIALES)

    if gauth.credentials is None:
        gauth.LocalWebserverAuth(port_numbers=[8092])
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()

    gauth.SaveCredentialsFile(CREDENCIALES)
    return GoogleDrive(gauth)

def subir_a_drive(ruta_archivo, id_folder=ID_CARPETA_DRIVE):
    drive = login_drive()
    archivo_drive = drive.CreateFile({
        "parents": [{"id": id_folder}],
        "title": os.path.basename(ruta_archivo)
    })
    archivo_drive.SetContentFile(ruta_archivo)
    archivo_drive.Upload()

    # 🔥 Hacemos el archivo visible a cualquiera con el link
    archivo_drive.InsertPermission({
        'type': 'anyone',
        'role': 'reader'
    })

    link = archivo_drive['alternateLink']
    print(f"📤 Subido a Drive: {ruta_archivo}")
    print(f"🔗 Link del PDF: {link}")

    return link
# ================================
# LEER GOOGLE SHEETS
# ================================
from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = "key_sheets.json"
SPREADSHEET_ID = "1o0SG78GLAPsTK-jGKVkNf2iOHyX9ycJMP-9PzlZYMro"
HOJA = "Hoja 1"


def leer_google_sheet_columna(columna_buscada):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    hoja = service.spreadsheets()

    result = hoja.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:Z"
    ).execute()

    valores = result.get("values", [])

    if not valores:
        print("❌ No se encontró información en Google Sheets")
        return []

    encabezados = valores[0]
    filas = valores[1:]

    try:
        index_col = encabezados.index(columna_buscada)
    except ValueError:
        print(f"❌ No existe la columna '{columna_buscada}'")
        print("Columnas disponibles:", encabezados)
        return []

    lista = []
    for fila in filas:
        if len(fila) > index_col and fila[index_col].strip():
            lista.append(fila[index_col].strip())

    print(f"📄 Google Sheet cargado: {len(lista)} clientes encontrados")
    return lista

def escribir_link_en_sheet(cliente, link):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service_gs = build('sheets', 'v4', credentials=creds)
    hoja = service_gs.spreadsheets()

    # Leemos A y B para ubicar la fila del cliente
    result = hoja.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:B"
    ).execute()

    valores = result.get("values", [])

    # Buscar fila donde está el cliente
    for i, fila in enumerate(valores):
        if len(fila) > 0 and fila[0] == cliente:
            fila_destino = i + 1  # Sheets index
            break
    else:
        print(f"⚠ Cliente {cliente} no encontrado en Sheets.")
        return

    rango = f"{HOJA}!B{fila_destino}"

    # 🔥 GUARDAR ENLACE COMO HYPERLINK
    body = {"values": [[f'=HYPERLINK("{link}"; "Ver PDF")']]}

    hoja.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rango,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

    print(f"📝 Enlace agregado en Google Sheets (Cliente: {cliente})")

def escribir_estado_deuda(cliente, estado):
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service_gs = build('sheets', 'v4', credentials=creds)
    hoja = service_gs.spreadsheets()

    # Leer encabezados y todas las filas
    result = hoja.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{HOJA}!A:Z"
    ).execute()

    valores = result.get("values", [])

    if not valores:
        print("⚠ No se pudo leer el Google Sheet.")
        return

    encabezados = valores[0]      # Primera fila
    filas = valores[1:]           # Resto de filas

    # Buscar la columna que se llama EXACTAMENTE "Estado de Cuenta"
    try:
        col_estado = encabezados.index("Estado de Cuenta")
    except ValueError:
        print("⚠ No existe la columna 'Estado de Cuenta' en Google Sheets.")
        print("Columnas encontradas:", encabezados)
        return

    # Buscar la fila del cliente
    fila_destino = None
    for i, fila in enumerate(filas, start=2):   # start=2 porque fila 1 es encabezado
        if len(fila) > 0 and fila[0] == cliente:
            fila_destino = i
            break

    if fila_destino is None:
        print(f"⚠ Cliente {cliente} no encontrado en Google Sheets.")
        return

    # Obtener el rango donde se debe escribir (EJ: C5, D7, etc.)
    columna_letra = chr(ord('A') + col_estado)
    rango = f"{HOJA}!{columna_letra}{fila_destino}"

    body = {"values": [[estado]]}

    hoja.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rango,
        valueInputOption="RAW",
        body=body
    ).execute()

    print(f"📌 Estado registrado en Sheets (Cliente {cliente} → {estado})")


# ================================
# CONFIGURACIÓN SELENIUM
# ================================
CARPETA_DESCARGAS = r"C:\Users\antho\Downloads\Facturas"
os.makedirs(CARPETA_DESCARGAS, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-gpu")
options.add_argument("--disable-infobars")

prefs = {
    "download.default_directory": CARPETA_DESCARGAS,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 30)


def click_ui5(element):
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", element)

    try:
        driver.execute_script("""
            try {
                var ctrl = sap.ui.getCore().byId(arguments[0].id.replace("-inner",""));
                if (ctrl) ctrl.firePress();
            } catch(e) {}
        """, element)
    except:
        pass


# ================================
# FUNCIÓN ORIGINAL — SOLO RENOMBRA Y SUBE
# ================================
def esperar_y_renombrar_pdf(cliente):
    print("⏳ Esperando descarga del PDF...")
    inicio = time.time()
    TIMEOUT = 60

    archivos_iniciales = set(os.listdir(CARPETA_DESCARGAS))

    while time.time() - inicio < TIMEOUT:
        archivos = os.listdir(CARPETA_DESCARGAS)

        en_descarga = [f for f in archivos if f.endswith(".crdownload")]
        pdfs = [f for f in archivos if f.lower().endswith(".pdf")]
        pdfs_nuevos = [f for f in pdfs if f not in archivos_iniciales]

        if not en_descarga and pdfs_nuevos:
            pdfs_nuevos.sort(key=lambda x: os.path.getmtime(os.path.join(CARPETA_DESCARGAS, x)))
            pdf_final = pdfs_nuevos[-1]

            fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            nombre_nuevo = f"{cliente}_{fecha}.pdf"

            ruta_pdf = os.path.join(CARPETA_DESCARGAS, pdf_final)
            ruta_final = os.path.join(CARPETA_DESCARGAS, nombre_nuevo)

            os.rename(ruta_pdf, ruta_final)
            print(f"📄 PDF guardado como {nombre_nuevo}")

            # Subir a Drive y obtener link
            link_pdf = subir_a_drive(ruta_final)

            # Guardar link en Google Sheets (columna B)
            escribir_link_en_sheet(cliente, link_pdf)

            return True

        time.sleep(1)

    print("❌ No se detectó PDF descargado")
    return False


# ================================
# CARGAR CLIENTES
# ================================
lista_clientes = leer_google_sheet_columna("Numero de Cliente")


# ================================
# PROCESO PRINCIPAL
# ================================
driver.get("https://www.metrogas.com.ar/consulta-y-paga-tu-saldo/")
wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
print("🚀 Página abierta")

iframe = wait.until(
    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src,'saldos.micuenta.metrogas.com.ar')]"))
)
driver.switch_to.frame(iframe)

for cliente in lista_clientes:

    print("\n==============================")
    print(f"➡ Procesando cliente: {cliente}")
    print("==============================")

    input_cliente = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--inputCustNumId-inner"))
    )
    input_cliente.clear()
    input_cliente.send_keys(cliente)

    driver.execute_script("""
        var input = document.getElementById('container-ovWebAbierta---Main--inputCustNumId-inner');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    time.sleep(1)

    boton_buscar = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--idButtonSearch-inner"))
    )
    click_ui5(boton_buscar)
    print("🔍 Buscar clickeado")

    try:
        wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "sapUiLocalBusyIndicator")))
    except:
        pass

    print("🧾 Buscando icono PDF…")
    XPATH_PDF = "//button[contains(@id,'idTableDebts') and contains(@id,'-0')]"
    XPATH_FRENTE = "//span[contains(text(), 'Frente')]/ancestor::div[contains(@id,'tile')]"

    try:
        boton_pdf = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PDF))
        )
        print("🧾 Icono PDF encontrado ✔")

        # ⬅ Cliente CON DEUDA
        escribir_estado_deuda(cliente, "Con deudas")

        click_ui5(boton_pdf)

    except Exception:
        print("⚠ No se encontró icono PDF. Abriendo mosaico 'Frente'…")

        # ⬅ Cliente SIN DEUDA
        escribir_estado_deuda(cliente, "Sin deudas")

        boton_frente = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_FRENTE))
        )
        click_ui5(boton_frente)


    time.sleep(2)

    print("⬇ Botón DESCARGAR…")
    boton_descargar = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//bdi[contains(text(),'Descargar')]/ancestor::button"))
    )
    click_ui5(boton_descargar)

    esperar_y_renombrar_pdf(cliente)

    print("🔄 Nueva consulta...")
    boton_nueva = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//bdi[contains(text(), 'Nueva consulta')]/ancestor::button"))
    )
    click_ui5(boton_nueva)

    time.sleep(1)

# ================================
# LIMPIEZA FINAL — BORRAR CARPETA COMPLETA
# ================================
print("\n🧹 Limpiando carpeta de descargas...")

import shutil

if os.path.exists(CARPETA_DESCARGAS):
    try:
        shutil.rmtree(CARPETA_DESCARGAS)
        print(f"🗑 Carpeta eliminada: {CARPETA_DESCARGAS}")
    except Exception as e:
        print(f"⚠ Error eliminando la carpeta: {e}")
else:
    print("⚠ Carpeta no existía, creando una nueva.")

# Recrear la carpeta vacía
try:
    os.makedirs(CARPETA_DESCARGAS, exist_ok=True)
    print(f"📁 Carpeta recreada: {CARPETA_DESCARGAS}")
except Exception as e:
    print(f"⚠ Error creando la carpeta: {e}")


print("\n🎉 PROCESO COMPLETADO")
driver.quit()
