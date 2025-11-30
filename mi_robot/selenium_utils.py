import os
import time
from datetime import datetime
from drive_utils import subir_a_drive
from sheets_utils import escribir_link_en_sheet

def click_ui5(driver, element):
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


def esperar_y_renombrar_pdf(driver, carpeta, cliente):
    inicio = time.time()
    archivos_iniciales = set(os.listdir(carpeta))

    while time.time() - inicio < 60:
        pdfs = [f for f in os.listdir(carpeta) if f.endswith(".pdf")]
        nuevos = [p for p in pdfs if p not in archivos_iniciales]

        if nuevos:
            archivo = max(nuevos, key=lambda x: os.path.getmtime(os.path.join(carpeta, x)))
            fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            nuevo_nombre = f"{cliente}_{fecha}.pdf"

            original = os.path.join(carpeta, archivo)
            destino = os.path.join(carpeta, nuevo_nombre)
            os.rename(original, destino)

            link = subir_a_drive(destino)
            escribir_link_en_sheet(cliente, link)
            return True
        time.sleep(1)

    return False
