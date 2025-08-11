import os
import time
import shutil
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from concurrent.futures import ThreadPoolExecutor, as_completed
import re

# Directorios de origen y destino
source_folder = r"H:\Mi unidad\REGISTROS PEDAGOGICOS - WEB\Pendientes"
destination_folder = r"H:\Mi unidad\REGISTROS PEDAGOGICOS - WEB\Generados"

# Obtener la carpeta de descargas del usuario
download_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# Configurar opciones de Chrome para modo headless y definir la carpeta de descargas
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1366,738")
prefs = {"download.default_directory": download_folder}
chrome_options.add_experimental_option("prefs", prefs)

def process_file(file_path, retry=3):
    for attempt in range(retry):
        driver = webdriver.Chrome(options=chrome_options)
        try:
            print(f"[INFO] Procesando archivo: {file_path}")

            # Aquí comienza el proceso con Selenium
            driver.get("https://vbaprotect.com/vbapro/main/")

            upload_field = driver.find_element(By.NAME, "upload_file")
            upload_field.send_keys(file_path)  # Enviar directamente la ruta del archivo
            print(f"[INFO] Archivo subido: {file_path}")

            submit_button = driver.find_element(By.CSS_SELECTOR, ".btn")
            submit_button.click()
            print(f"[INFO] Botón de enviar clicado para archivo: {file_path}")

            # Esperar hasta que el botón de creación esté visible y se pueda hacer clic en él
            create_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.NAME, "create_btn"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", create_btn)  # Desplazar la vista hacia el botón
            driver.execute_script("arguments[0].click();", create_btn)  # Hacer clic usando JavaScript
            print(f"[INFO] Botón de creación clicado para archivo: {file_path}")

            # Esperar hasta que el archivo se procese y el botón de descarga aparezca
            download_btn_locator = (By.XPATH, "/html/body/div[2]/div[2]/form/div[1]/h6/button")
            download_btn = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(download_btn_locator)
            )
            
            # Desplazar la vista hacia el botón y hacer clic en él
            driver.execute_script("arguments[0].scrollIntoView(true);", download_btn)
            driver.execute_script("arguments[0].click();", download_btn)
            print(f"[INFO] Botón de descarga clicado para archivo: {file_path}")

            # Esperar a que el archivo se descargue completamente
            downloaded_file = ""
            while not downloaded_file or downloaded_file.endswith('.crdownload'):
                time.sleep(1)
                downloaded_file = max([os.path.join(download_folder, f) for f in os.listdir(download_folder)], key=os.path.getctime)

            print(f"[INFO] Descarga completa para el archivo: {file_path}")

            # Obtener el nombre base esperado del archivo descargado
            base_filename = re.sub(r'_[a-zA-Z0-9]+_ptd', '', os.path.basename(downloaded_file))
            base_filename = re.sub(r'_ptd', '', base_filename)

            # Mover el archivo descargado a la carpeta de destino correspondiente
            final_destination_path = os.path.join(destination_folder, base_filename)
            shutil.move(downloaded_file, final_destination_path)
            print(f"[INFO] Archivo movido a: {final_destination_path}")
            break  # Salir del bucle si el procesamiento es exitoso

        except Exception as e:
            print(f"[ERROR] Error procesando el archivo {file_path} (Intento {attempt + 1}/{retry}): {e}")
        finally:
            driver.quit()
            print(f"[INFO] Navegador cerrado para archivo: {file_path}")
    else:
        print(f"[ERROR] El archivo {file_path} no se pudo procesar después de {retry} intentos.")

def process_folder(professor_folder_path):
    tasks = []
    for filename in os.listdir(professor_folder_path):
        if filename.endswith(".xlsm") and not filename.startswith("~$"):
            file_path = os.path.join(professor_folder_path, filename)
            tasks.append(file_path)

    # Ejecutar los procesos en paralelo, limitando a 5 ejecuciones simultáneas
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(process_file, task) for task in tasks]
        for future in as_completed(futures):
            future.result()

    # Eliminar la carpeta después de procesar todos los archivos
    shutil.rmtree(professor_folder_path)
    print(f"[INFO] Carpeta eliminada: {professor_folder_path}")

# Crear una lista de carpetas a procesar
folders_to_process = []

# Recorrer las subcarpetas en la carpeta de origen
for root, dirs, files in os.walk(source_folder):
    # Verificar que no es la carpeta raíz
    if root != source_folder:
        folders_to_process.append(root)

# Procesar las carpetas con un máximo de 5 ejecuciones simultáneas
with ThreadPoolExecutor(max_workers=5) as folder_executor:
    folder_futures = [folder_executor.submit(process_folder, folder) for folder in folders_to_process]
    for folder_future in as_completed(folder_futures):
        folder_future.result()

print("[INFO] Proceso completado, no hay carpetas para procesar.")

# Ejecutar el script web.py al finalizar todos los procesos
def ejecutar_web_script():
    script_path = "G:\\Mi unidad\\PROYECTOS PARA VENTAS\\Pagina - EduFacil\\Configurar servidor para interaccion con pagina\\04 - Generador de enlaces - descargas.py"
    try:
        subprocess.run(["python", script_path], check=True)
        print("Script web.py ejecutado exitosamente.")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el script web.py: {e}")

ejecutar_web_script()
