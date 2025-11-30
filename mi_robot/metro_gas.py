import os
import shutil
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from sheets_utils import leer_google_sheet_columna, escribir_estado_deuda
from selenium_utils import click_ui5, esperar_y_renombrar_pdf


CARPETA_DESCARGAS = r"C:\Users\antho\Downloads\Facturas"


def procesar_metro_gas():

    if not os.path.exists(CARPETA_DESCARGAS):
        os.makedirs(CARPETA_DESCARGAS)

    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": CARPETA_DESCARGAS,
        "download.prompt_for_download": False
    })

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    wait = WebDriverWait(driver, 30)

    lista_clientes = leer_google_sheet_columna("Numero de Cliente")

    driver.get("https://www.metrogas.com.ar/consulta-y-paga-tu-saldo/")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    iframe = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[contains(@src,'saldos.micuenta.metrogas.com.ar')]")
    ))
    driver.switch_to.frame(iframe)

    for cliente in lista_clientes:
        input_cliente = wait.until(EC.element_to_be_clickable(
            (By.ID, "container-ovWebAbierta---Main--inputCustNumId-inner")
        ))
        input_cliente.clear()
        input_cliente.send_keys(cliente)

        time.sleep(1)

        boton_buscar = wait.until(EC.element_to_be_clickable(
            (By.ID, "container-ovWebAbierta---Main--idButtonSearch-inner")
        ))
        click_ui5(driver, boton_buscar)

        try:
            boton_pdf = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@id,'idTableDebts') and contains(@id,'-0')]")
            ))
            escribir_estado_deuda(cliente, "Con deudas")
            click_ui5(driver, boton_pdf)

        except:
            escribir_estado_deuda(cliente, "Sin deudas")
            mosaico = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(), 'Frente')]/ancestor::div[contains(@id,'tile')]")
            ))
            click_ui5(driver, mosaico)

        time.sleep(1)

        boton_descargar = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//bdi[contains(text(),'Descargar')]/ancestor::button")
        ))
        click_ui5(driver, boton_descargar)

        esperar_y_renombrar_pdf(driver, CARPETA_DESCARGAS, cliente)

        boton_nueva = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//bdi[contains(text(),'Nueva consulta')]/ancestor::button")
        ))
        click_ui5(driver, boton_nueva)

    shutil.rmtree(CARPETA_DESCARGAS)
    os.makedirs(CARPETA_DESCARGAS)

    driver.quit()
