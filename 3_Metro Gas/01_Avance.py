from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import time
import os

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
options.add_argument("--disable-infobars")
options.add_argument("--start-maximized")

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
# FUNCIONES
# =============================================

def click_sap_ui5(driver, by, value, timeout=25):
    """Hace un click seguro sobre controles SAP UI5."""
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, value))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", el)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", el)
        
        try:
            driver.execute_script(f"""
                try {{
                    sap.ui.getCore().byId("{value.replace('-inner','')}").firePress();
                }} catch(e) {{}}
            """)
        except:
            pass
        
        return True
    except Exception as e:
        print(f"⚠ Error SAP UI5 click en {value}: {e}")
        return False

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

    # Evento UI5
    driver.execute_script("""
        var input = document.getElementById('container-ovWebAbierta---Main--inputCustNumId-inner');
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    """)

    time.sleep(1)

    # CLIC EN "BUSCAR"
    print("🔍 Click en botón Buscar...")
    if click_sap_ui5(driver, By.ID, "container-ovWebAbierta---Main--idButtonSearch-inner"):
        print("✓ Click Buscar OK")
    elif click_sap_ui5(driver, By.ID, "container-ovWebAbierta---Main--idButtonSearch"):
        print("✓ Click Buscar OK (wrapper)")
    else:
        raise Exception("Botón Buscar no respondió")

    # Esperamos datos
    try:
        wait.until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "sapUiLocalBusyIndicator"))
        )
    except:
        pass

    print("✓ Datos cargados correctamente")

    # CLIC EN "ÚLTIMA FACTURA"
    print("📄 Abriendo 'Última Factura'...")

    TAB_ID = "container-ovWebAbierta---Main--idTab1"
    TAB_INNER = "container-ovWebAbierta---Main--idTab1-tab"

    if click_sap_ui5(driver, By.ID, TAB_INNER):
        print("✓ Tab 'Última Factura' clickeado")
        time.sleep(2)
    elif click_sap_ui5(driver, By.ID, TAB_ID):
        print("✓ Tab 'Última Factura' clickeado (wrapper)")
        time.sleep(2)
    else:
        try:
            el = wait.until(EC.element_to_be_clickable((By.ID, TAB_ID)))
            driver.execute_script("arguments[0].click();", el)
            print("✓ Tab clickeado (backup)")
            time.sleep(2)
        except Exception as e:
            raise e

    # Esperar carga del tab
    try:
        wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "sapUiLocalBusyIndicator")))
    except:
        pass

    print("✓ 'Última Factura' abierta")
    
    # Tomar screenshot
    driver.save_screenshot(os.path.join(CARPETA_DESCARGAS, "ultima_factura.png"))
    print("📸 Screenshot guardado")

    print("\n" + "="*50)
    print("✅ PROCESO COMPLETADO")
    print("   Se detuvo en 'Última Factura'")
    print(f"📸 Screenshot en: {CARPETA_DESCARGAS}")
    print("="*50)

except Exception as e:
    print(f"\n❌ ERROR: {e}")
    driver.save_screenshot(os.path.join(CARPETA_DESCARGAS, "error.png"))

finally:
    print("\n⏸ Navegador abierto 60 segundos...")
    print("   (Podés ver la pantalla de 'Última Factura')")
    time.sleep(60)
    driver.quit()
    print("✓ Navegador cerrado")