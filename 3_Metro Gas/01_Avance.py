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

CREDENCIALES = "credentials_modulo.json"
ID_CARPETA_DRIVE = "1BCb9JXw17Eix4Ux2BIzN1ok50lgFLb_H"   # ← tu carpeta de Drive

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
    print(f"📤 Subido a Drive: {ruta_archivo}")


# =============================================
# CONFIGURACIÓN
# =============================================

EXCEL_PATH = "Metrogas.xlsx"
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

# =============================================
# CLICK UI5 LOW LEVEL
# =============================================

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

# =============================================
# FUNCIÓN ESPERAR DESCARGA
# =============================================

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

            # 📤 SUBIR A GOOGLE DRIVE
            subir_a_drive(ruta_final)

            return True

        time.sleep(1)

    print("❌ No se detectó PDF descargado")
    return False

# =============================================
# LEER EXCEL
# =============================================

df = pd.read_excel(EXCEL_PATH)

# Asegurar coincidencia exacta del nombre de columna
columna_correcta = [col for col in df.columns if col.strip().lower() == "numero de cliente"]
if not columna_correcta:
    print("❌ No se encontró la columna 'Numero de Cliente' en el Excel")
    print("Columnas disponibles:", list(df.columns))
    exit()

col = columna_correcta[0]
lista_clientes = df[col].astype(str).tolist()

# =============================================
# ABRIR WEB
# =============================================

driver.get("https://www.metrogas.com.ar/consulta-y-paga-tu-saldo/")
wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
print("🚀 Página abierta")

# Entrar al iframe
iframe = wait.until(
    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src,'saldos.micuenta.metrogas.com.ar')]"))
)
driver.switch_to.frame(iframe)

# =============================================
# INICIO DEL LOOP
# =============================================

for cliente in lista_clientes:

    print("\n==============================")
    print(f"➡ Procesando cliente: {cliente}")
    print("==============================")

    # INPUT CLIENTE
    input_cliente = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--inputCustNumId-inner"))
    )
    input_cliente.clear()
    input_cliente.send_keys(cliente)

    # Disparar eventos UI5
    driver.execute_script("""
        var input = document.getElementById('container-ovWebAbierta---Main--inputCustNumId-inner');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    time.sleep(1)

    # BOTÓN BUSCAR
    boton_buscar = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--idButtonSearch-inner"))
    )
    click_ui5(boton_buscar)
    print("🔍 Buscar clickeado")

    # Esperar carga
    try:
        wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "sapUiLocalBusyIndicator")))
    except:
        pass

    # BOTÓN ICONO PDF EN TABLA
    print("🧾 Click en icono PDF de la tabla…")
    XPATH_PDF = "//button[contains(@id,'idTableDebts') and contains(@id,'-0')]"

    boton_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_PDF)))
    click_ui5(boton_pdf)

    time.sleep(2)

    # BOTÓN DESCARGAR
    print("⬇ Botón DESCARGAR…")
    XPATH_DESCARGAR = "//bdi[contains(text(),'Descargar')]/ancestor::button"

    boton_descargar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_DESCARGAR)))
    click_ui5(boton_descargar)

    # ESPERAR Y RENOMBRAR PDF AUTOMÁTICAMENTE
    esperar_y_renombrar_pdf(cliente)

    # BOTÓN "NUEVA CONSULTA"
    print("🔄 Volver a 'Nueva consulta'...")
    XPATH_NUEVA = "//bdi[contains(text(), 'Nueva consulta')]/ancestor::button"

    boton_nueva = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_NUEVA)))
    click_ui5(boton_nueva)

    time.sleep(1)

print("\n🎉 PROCESO COMPLETADO CON ÉXITO")
driver.quit()
