# Generador
# *************************
import os
import time
from http.client import HTTPException

import requests
import json
import gc
import os
import openpyxl
import xlsxwriter

import subprocess
from zipfile import ZipFile

# *************************

# Obfuscador
# *************************
import shutil
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
# *************************

# Super - Subir Enlaces
# *************************
# *************************

import uvicorn
from fastapi import FastAPI, status
from fastapi.responses import FileResponse, JSONResponse

app = FastAPI()

# Code Generador
# ***************************************************************************************************************
def obtener_compras():
    """Obtiene las compras desde la API."""
    print("[DEBUG] Iniciando obtención de compras desde la API...")
    url = "https://www.edufacil.net/Actualizacion/updates/API/API%20PYTHON/api.php"
    try:
        response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps({}))
        response.raise_for_status()
        datos = response.json()

        print("[DEBUG] Respuesta de la API:", datos)

        if isinstance(datos, dict) and "error" in datos:
            print("[DEBUG] Mensaje de la API:", datos["error"])
            return []  # Retorna una lista vacía si no hay compras
        elif isinstance(datos, list) and datos:
            # Filtrar registros duplicados basados en el campo 'id'
            registros_unicos = {registro['id']: registro for registro in datos}.values()
            registros_unicos = list(registros_unicos)
            print("[DEBUG] Datos obtenidos desde la API (sin duplicados):", registros_unicos)
            return registros_unicos
        else:
            print("[DEBUG] No se encontraron compras disponibles en la API.")
            return []

    except requests.exceptions.HTTPError as http_err:
        print(f"[ERROR] HTTP error occurred: {http_err}")
        return []
    except requests.exceptions.RequestException as req_err:
        print(f"[ERROR] Error en la solicitud: {req_err}")
        return []


def determinar_nivel_base(item_id):
    """Determina el nivel base según el ID del artículo."""
    if item_id == 1:
        return "PRIMARIA"
    elif item_id == 2:
        return "SECUNDARIA"
    elif item_id == 3:
        return "ALTERNATIVA"
    elif item_id == 4:
        return "INICIAL"
    else:
        return "DESCONOCIDO"

def actualizar_estado_api(ids):
    """Actualiza el estado del archivo en la base de datos utilizando la API actualizar_generado.php."""
    print("[DEBUG] Iniciando actualización de estado en la API...")
    url = "https://www.edufacil.net/Actualizacion/updates/API/API%20PYTHON/actualizar_generado.php"

    for archivo_id in ids:
        payload = {'id': archivo_id}
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload))
            response.raise_for_status()
            respuesta = response.json()
            if respuesta.get('success'):
                print(f"[DEBUG] Archivo con ID {archivo_id} marcado como generado.")
            else:
                print(
                    f"[ERROR] Al marcar el archivo con ID {archivo_id}: {respuesta.get('error', 'Error desconocido')}")
        except requests.exceptions.RequestException as e:
            print(f"[ERROR] Al marcar el archivo con ID {archivo_id}: {e}")

def reemplazar_espacios_por_guiones(texto):
    # Reemplaza uno o más espacios en blanco por un guion bajo
    return re.sub(r'\s+', '_', texto)

def extraer_macros(archivo_xlsm):
    """Extrae las macros de un archivo .xlsm utilizando vba_extract."""
    try:
        # Ejecutar la utilidad vba_extract para extraer las macros
        subprocess.run(['python3', 'vba_extract.py', archivo_xlsm], check=True)
        print(f"[DEBUG] Macros extraídas de {archivo_xlsm}.")
    except Exception as e:
        print(f"[ERROR] No se pudo extraer las macros: {e}")

def actualizar_modulo_vba_y_guardar(ruta_carpeta_niveles, registros):
    """Actualiza los datos en el módulo VBA y guarda los archivos Excel,
    luego llama a la macro ActualizarDatosCriticosEnHojas para que haga
    el resto de actualizaciones (p.ej. la hoja CARATULA).
    """
    ruta_macros = 'vbaProject.bin'

    print("[DEBUG] Iniciando actualización de módulos y guardado de archivos...")
    carpeta_generada = r"./generados"  # Ajusta la ruta si corresponde

    if not os.path.exists(carpeta_generada):
        os.makedirs(carpeta_generada)  # Se crea la carpeta
        print(f"[DEBUG] Carpeta creada: {carpeta_generada}")

    ids_procesados = []  # Almacenar los IDs de los registros procesados correctamente

    archivos_generados = []

    for registro in registros:
        unidad_educativa = registro['unidad_educativa']
        director = registro.get('director', '')
        profesor = registro['nombre_completo']
        nivel = determinar_nivel_base(registro['item_id'])

        # Verificar que el nivel tiene un archivo correspondiente
        archivo_nivel = os.path.join(ruta_carpeta_niveles, f"{nivel.lower()}.xlsm")
        if not os.path.exists(archivo_nivel):
            print(f"[ERROR] No se encontró el archivo para el nivel: {nivel} ({archivo_nivel})")
            continue

        print(f"[DEBUG] Procesando registro ID={registro['id']} - Nivel={nivel} - Prof={profesor}")

        # Extraer macros del archivo original
        #extraer_macros(archivo_nivel)

        try:
            # Crear carpeta para la unidad educativa
            unidad_educativa_new = unidad_educativa.lower()
            unidad_educativa_new = reemplazar_espacios_por_guiones(unidad_educativa_new)
            carpeta_unidad = os.path.join(carpeta_generada, unidad_educativa_new)
            if not os.path.exists(carpeta_unidad):
                os.makedirs(carpeta_unidad)
                print(f"[DEBUG] Carpeta creada: {carpeta_unidad}")

            nombre_archivo_excel = f"{registro['id']}_Tu Registro Pedagogico  {profesor.upper()}  {nivel}.xlsm"
            nombre_archivo_excel = reemplazar_espacios_por_guiones(nombre_archivo_excel)
            # ruta_nueva = os.path.join(carpeta_unidad, f"{registro['id']}_Tu Registro Pedagogico - {profesor.upper()} - {nivel}.xlsm")
            ruta_nueva = os.path.join(carpeta_unidad, nombre_archivo_excel)

            # Cargar el archivo Excel
#            workbook = xlsxwriter.Workbook(archivo_nivel, {'in_memory': True})
            workbook = openpyxl.load_workbook(archivo_nivel, data_only=True) #keep_links=False)
            print("[DEBUG] Archivo Excel cargado correctamente.")

            # Actualizar el módulo de datos críticos (simulación)
            # En lugar de un módulo VBA, simplemente actualizamos las celdas
            hoja_menu = workbook['MENU']
            hoja_menu['Q33'] = "1"  # Versión inicial
#            hoja_menu.write('Q33', '1')
            print("[DEBUG] Se estableció versión en la hoja MENU.")

            # Actualizar datos en la hoja (simulación de la macro)
            '''
            hoja_datos_criticos = workbook['DatosCriticos']  # Cambia el nombre según tu archivo
            hoja_datos_criticos['A1'] = unidad_educativa  # Ejemplo de actualización
            hoja_datos_criticos['A2'] = director
            hoja_datos_criticos['A3'] = profesor
            print("[DEBUG] Se actualizaron los datos críticos en la hoja.")
            '''

            # Verificar si el archivo de macros fue creado, version xlsxwriter
            '''
            if os.path.exists(ruta_macros):
                # Agregar las macros al archivo
                workbook.add_vba_project(ruta_macros)
                print("[INFO] Se agregaron las macros al archivo")
            else:
                print("[ERROR] El archivo de macros no fue encontrado. Este archivo presentara errores de formato")
            '''

            # Guardar el archivo final con el nuevo formato
            workbook.save(ruta_nueva)
            workbook.close()

            print(f"[DEBUG] Archivo guardado: {ruta_nueva}")

            archivos_generados.append(ruta_nueva)

            # Agregar el ID a la lista de procesados
            ids_procesados.append(registro['id'])
            print(f"[DEBUG] Generación exitosa para ID={registro['id']} -> {ruta_nueva}")

        except Exception as e:
            print(f"[ERROR] Al procesar el archivo para {profesor} - {nivel}: {e}")
            continue

    if archivos_generados:
        print("[INFO] Archivos generados exitosamente:")
        for archivo in archivos_generados:
            print("   ", archivo)
    else:
        print("[INFO] No se generaron archivos para ningún registro.")

    # Actualizar el estado en la API (solo si hay IDs que se procesaron correctamente)
    if ids_procesados:
        actualizar_estado_api(ids_procesados)
    else:
        print("[DEBUG] No hubo IDs para actualizar en la API.")

# ***************************************************************************************************************

# Code - Subir Enlaces
# ***************************************************************************************************************

# ***************************************************************************************************************

@ app.get("/download/{filename}")
def download_file(filename: str):  # async ?
    # Ruta del archivo que deseas descargar
    file_path = os.path.join("./generados/", filename)

    # Verificar si el archivo existe
    if os.path.exists(file_path):
        print("[INFO] El archivo a descargar existe...")
        return FileResponse(file_path)
    else:
        print("[DEBUG] El archivo a descargar NO existe...")
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
        #return JSONResponse(status_code=404)

@app.get("/generar_archivos_automaticamente")
async def generar_archivos_automaticamente():
    """Genera archivos automáticamente sin interfaz gráfica."""
    return_generador = "Ok"

    print("[DEBUG] Ejecutando generar_archivos_automaticamente()...")
    ruta_carpeta_niveles = r"./niveles"  # Ajusta la ruta si corresponde

    if not os.path.exists(ruta_carpeta_niveles):
        print(f"[ERROR] La carpeta de niveles no existe: {ruta_carpeta_niveles}")
        return_generador = "nOkE"  # return

    registros = obtener_compras()
    if not registros:
        print("[INFO] No hay registros para procesar (lista vacía).")
        return_generador = "nOkR"  # return
    else:
        actualizar_modulo_vba_y_guardar(ruta_carpeta_niveles, registros)

    return {"resultado": return_generador}  # return_generador  # print("[INFO] Proceso de generación completado.")
