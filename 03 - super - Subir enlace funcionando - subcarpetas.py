from __future__ import print_function
import os.path
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import requests
import json
import re
from google.auth.exceptions import RefreshError
import urllib3

# Deshabilitar advertencias de solicitudes inseguras
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Rutas de las credenciales y token
CREDENTIALS_FILE = r'G:\Mi unidad\PROYECTOS PARA VENTAS\Pagina - EduFacil\Configurar servidor para interaccion con pagina\credentials.json'
TOKEN_FILE = 'token.json'

# Los permisos que necesitamos (si solo necesitas lectura, usa 'readonly')
SCOPES = ['https://www.googleapis.com/auth/drive']

def authenticate_gdrive():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    try:
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            # Guardar las credenciales en el archivo token.json
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
    except RefreshError as e:
        print(f"[ERROR] Token inválido o expirado: {e}. Se requiere autenticación nuevamente.")
        if os.path.exists(TOKEN_FILE):
            os.remove(TOKEN_FILE)
            print(f"[DEBUG] Archivo {TOKEN_FILE} eliminado. Se pedirá autenticación nuevamente.")
        # Intentar autenticación nuevamente
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def list_files_and_folders_in_folder(service, folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(q=query).execute()
    items = results.get('files', [])
    return items

def process_folder(service, folder_id):
    items = list_files_and_folders_in_folder(service, folder_id)
    for item in items:
        if item['mimeType'] == 'application/vnd.google-apps.folder':
            print(f"[DEBUG] Procesando subcarpeta: {item['name']} (ID: {item['id']})")
            # Procesar recursivamente la subcarpeta
            process_folder(service, item['id'])
        else:
            print(f"[DEBUG] Procesando archivo: {item['name']}")
            process_file(service, item)

def process_file(service, file):
    try:
        file_id = extract_id_from_filename(file['name'])
        if file_id is not None:
            if check_file_exists(file_id):
                print(f"File with ID {file_id} already has a download link. Skipping...")
                return

            make_file_public(service, file['id'])
            download_link = get_file_download_link(file['id'])
            print(f"File: {file['name']} - Download Link: {download_link}")

            # Envía el enlace de descarga a la API
            send_download_link_to_api(file_id, file['name'], download_link, "Descripción del archivo")
            
            # Actualiza el campo archivo_generado en la base de datos
            update_file_generated(file_id)
        else:
            print(f"Failed to extract ID from filename: {file['name']}")
    except Exception as e:
        print(f"Error processing file {file['name']}: {e}")

def make_file_public(service, file_id):
    permission = {
        'type': 'anyone',
        'role': 'reader',
    }
    service.permissions().create(
        fileId=file_id,
        body=permission,
    ).execute()

def get_file_download_link(file_id):
    return f"https://drive.google.com/uc?export=download&id={file_id}"

def extract_id_from_filename(filename):
    match = re.match(r"(\d+)_", filename)
    if match:
        return int(match.group(1))
    return None

def check_file_exists(id):
    url = "https://edufacil.net/Actualizacion/updates/Tienda/verificar_archivo.php"  # Reemplaza con la URL de tu servidor
    data = {"id": id}
    headers = {'Content-Type': 'application/json'}
    try:
        response = requests.post(url, data=json.dumps(data), headers=headers, verify=False)  # Añadir verify=False
        if response.status_code == 200:
            result = response.json()
            return result['enlace_descarga'] is not None
        else:
            print(f"Failed to verify file existence for ID {id}. Error: {response.text}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Error al verificar el archivo con ID {id}: {e}")
        return False

def send_download_link_to_api(id, file_name, download_link, descripcion):
    url = "https://edufacil.net/Actualizacion/updates/Tienda/registrar_archivo.php"  # Reemplaza con la URL de tu servidor
    data = {
        "id": id,
        "descripcion": descripcion,
        "enlace_descarga": download_link    
    }
    headers = {'Content-Type': 'application/json'}
    
    # Validar que el enlace de descarga no esté vacío
    if not download_link:
        print(f"[ERROR] El enlace de descarga está vacío para el archivo {file_name}. No se enviará a la API.")
        return

    # Imprimir datos enviados para depuración
    print(f"[DEBUG] Datos enviados a la API: {data}")
    
    try:
        response = requests.post(url, data=json.dumps(data), headers=headers, verify=False)
        if response.status_code == 200:
            print(f"Successfully updated data for {file_name}")
        else:
            print(f"Failed to update data for {file_name}. Error: {response.status_code} - {response.text}")
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Error al enviar el enlace para {file_name}: {e}")

def update_file_generated(id):
    url = "https://edufacil.net/Actualizacion/updates/API/API%20PYTHON/actualizar_generado.php"
    data = {"id": id}
    headers = {'Content-Type': 'application/json'}

    try:
        response = requests.post(url, data=json.dumps(data), headers=headers, verify=False)
        if response.status_code == 200:
            print(f"[SUCCESS] Registro con ID {id} actualizado correctamente en la base de datos.")
        else:
            print(f"[ERROR] No se pudo actualizar el registro con ID {id}. Error: {response.status_code} - {response.text}")
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Error al realizar la solicitud a la API para ID {id}: {e}")

def main():
    print("[DEBUG] Generando enlaces.py se ha iniciado.")
    creds = authenticate_gdrive()
    service = build('drive', 'v3', credentials=creds)
    
    folder_id = '1-DXNbjbUuYw5RwZiNhS35urWu1rdupmp'
    process_folder(service, folder_id)

    print("[DEBUG] web.py se ha completado.")

if __name__ == '__main__':
    main()
