from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time, os

# =============================================
# CONFIGURACIÓN
# =============================================

NUMERO_CLIENTE = "10206069501"
CARPETA_DESCARGAS = r"C:\Users\antho\Downloads\Facturas"
os.makedirs(CARPETA_DESCARGAS, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")
options.add_argument("--disable-gpu")
options.add_argument("--start-maximized")
options.add_argument("--disable-infobars")

prefs = {
    "download.default_directory": CARPETA_DESCARGAS,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
    "download.directory_upgrade": True,
    "profile.default_content_settings.popups": 0,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 30)

# =============================================
# CLICK SAP UI5
# =============================================

def click_ui5(element):
    """Fuerza el click UI5 a bajo nivel."""
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
# PROCESO PRINCIPAL
# =============================================

try:
    print("🚀 Iniciando")
    driver.get("https://www.metrogas.com.ar/consulta-y-paga-tu-saldo/")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    print("✓ Página cargada")

    # IFRAME
    iframe = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@src,'saldos.micuenta.metrogas.com.ar')]")
        )
    )
    driver.switch_to.frame(iframe)
    print("✓ Dentro del iframe")

    # INPUT CLIENTE
    input_cliente = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "container-ovWebAbierta---Main--inputCustNumId-inner")
        )
    )
    input_cliente.clear()
    input_cliente.send_keys(NUMERO_CLIENTE)
    print(f"✓ Número ingresado: {NUMERO_CLIENTE}")

    driver.execute_script("""
        var input = document.getElementById('container-ovWebAbierta---Main--inputCustNumId-inner');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    time.sleep(1)

    # BOTÓN BUSCAR
    print("🔍 Click en botón Buscar...")

    boton_buscar = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--idButtonSearch-inner"))
    )
    click_ui5(boton_buscar)
    print("✓ Click Buscar OK")

    # ESPERA DE CARGA
    try:
        wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "sapUiLocalBusyIndicator")))
    except:
        pass

    print("✓ Datos cargados correctamente")

    # BOTÓN PDF
    print("🖨 Click en botón PDF…")

    boton_pdf = wait.until(
        EC.element_to_be_clickable((By.ID, "container-ovWebAbierta---Main--idDebtPDFButton-inner"))
    )
    click_ui5(boton_pdf)
    print("✓ PDF clickeado (inner)")

    time.sleep(2)

    # BOTÓN DESCARGAR
    print("⬇ Buscando botón 'Descargar' (ID variable)…")

    XPATH_DESCARGAR = "//bdi[contains(text(), 'Descargar')]/ancestor::button"
    boton_descargar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_DESCARGAR)))

    click_ui5(boton_descargar)
    print("✅ Botón DESCARGAR presionado")

    # =============================================
    # CAPTURA NUEVA PESTAÑA
    # =============================================

    print("📄 Esperando nueva pestaña con PDF…")
    time.sleep(2)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        print("📄 PDF abierto en nueva pestaña")
    else:
        print("❗ No se abrió pestaña nueva, pero el PDF puede estar descargándose")

except Exception as e:
    print(f"\n❌ ERROR: {e}")

finally:
    print("\n⏸ Dejando el navegador abierto 30s…")
    time.sleep(30)
    driver.quit()
    print("✓ Navegador cerrado")
