import os
import openpyxl
import gc

def actualizar_modulo_vba_y_guardar(ruta_carpeta_niveles, registros):
    """Actualiza los datos en el módulo VBA y guarda los archivos Excel,
    luego llama a la macro ActualizarDatosCriticosEnHojas para que haga
    el resto de actualizaciones (p.ej. la hoja CARATULA).
    """
    print("[DEBUG] Iniciando actualización de módulos y guardado de archivos...")
    carpeta_generada = r"./root/generados"  # Ajusta la ruta si corresponde

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

        try:
            # Crear carpeta para la unidad educativa
            carpeta_unidad = os.path.join(carpeta_generada, unidad_educativa)
            if not os.path.exists(carpeta_unidad):
                os.makedirs(carpeta_unidad)
                print(f"[DEBUG] Carpeta creada: {carpeta_unidad}")

            ruta_nueva = os.path.join(carpeta_unidad, f"{registro['id']}_Tu Registro Pedagogico - {profesor.upper()} - {nivel}.xlsm")

            # Cargar el archivo Excel
            workbook = openpyxl.load_workbook(archivo_nivel, keep_links=False)
            print("[DEBUG] Archivo Excel cargado correctamente.")

            # Actualizar el módulo de datos críticos (simulación)
            # En lugar de un módulo VBA, simplemente actualizamos las celdas
            hoja_menu = workbook['MENU']
            hoja_menu['Q33'] = "1"  # Versión inicial
            print("[DEBUG] Se estableció versión en la hoja MENU.")

            # Actualizar datos en la hoja (simulación de la macro)
            hoja_datos_criticos = workbook['DatosCriticos']  # Cambia el nombre según tu archivo
            hoja_datos_criticos['A1'] = unidad_educativa  # Ejemplo de actualización
            hoja_datos_criticos['A2'] = director
            hoja_datos_criticos['A3'] = profesor
            print("[DEBUG] Se actualizaron los datos críticos en la hoja.")

            # Guardar el archivo final con el nuevo formato
            workbook.save(ruta_nueva)
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

# Original

def actualizar_modulo_vba_y_guardar(ruta_carpeta_niveles, registros):
    """Actualiza los datos en el módulo VBA y guarda los archivos Excel,
    luego llama a la macro ActualizarDatosCriticosEnHojas para que haga
    el resto de actualizaciones (p.ej. la hoja CARATULA).
    """
    print("[DEBUG] Iniciando actualización de módulos VBA y guardado de archivos...")
    excel = None
    carpeta_generada = r"./root/generados"  # Ajusta la ruta si corresponde

    '''
    if not os.path.exists(carpeta_generada):
        os.makedirs(carpeta_generada)
        print(f"[DEBUG] Carpeta creada: {carpeta_generada}")
    '''

    ids_procesados = []  # Almacenar los IDs de los registros procesados correctamente

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        print("[DEBUG] Excel.Application iniciado correctamente.")

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

            try:
                # Crear carpeta para la unidad educativa
                carpeta_unidad = os.path.join(carpeta_generada, unidad_educativa)
                if not os.path.exists(carpeta_unidad):
                    os.makedirs(carpeta_unidad)
                    print(f"[DEBUG] Carpeta creada: {carpeta_unidad}")

                ruta_copia_temporal = os.path.join(carpeta_unidad, "temp_copy.xlsm")
                workbook_nivel = excel.Workbooks.Open(archivo_nivel)
                workbook_nivel.SaveCopyAs(ruta_copia_temporal)
                workbook_nivel.Close(SaveChanges=False)
                print("[DEBUG] Se generó copia temporal del archivo base.")

                workbook = excel.Workbooks.Open(ruta_copia_temporal)

                # 1) Buscar el módulo "ModuloDatosCriticos"
                vb_project = workbook.VBProject
                modulo = None
                for vb_component in vb_project.VBComponents:
                    if vb_component.Name == "ModuloDatosCriticos":
                        modulo = vb_component
                        break

                if not modulo:
                    workbook.Close(SaveChanges=False)
                    raise ValueError("[ERROR] No se encontró el módulo 'ModuloDatosCriticos'.")

                # 2) Insertar nuevo código en el módulo "ModuloDatosCriticos"
                nuevo_codigo = f"""
' Module: ModuloDatosCriticos

Public Const UnidadEducativa As String = "{unidad_educativa}"
Public Const Director As String = "{director}"
Public Const Profesor As String = "{profesor}"
"""
                modulo.CodeModule.DeleteLines(1, modulo.CodeModule.CountOfLines)
                modulo.CodeModule.AddFromString(nuevo_codigo.strip())
                print("[DEBUG] Se actualizó el módulo VBA con los nuevos datos.")

                # 3) Hoja MENU: asignar valor a Q33 (Versión)
                hoja_menu = workbook.Sheets("MENU")
                hoja_menu.Range("Q33").Value = "1"  # Versión inicial
                print("[DEBUG] Se estableció versión en la hoja MENU.")

                # ---------------------------------------------------------
                #   BLOQUE ELIMINADO (CARATULA) – Usaremos la macro en su lugar
                # ---------------------------------------------------------

                # 4) Llamar la macro ActualizarDatosCriticosEnHojas
                #    Asegúrate de que la macro es 'Public Sub ActualizarDatosCriticosEnHojas'
                #    en un módulo estándar del workbook, y que no es Private Sub.
                try:
                    macro_name = f"'{workbook.Name}'!ActualizarDatosCriticosEnHojas"
                    excel.Run(macro_name)
                    print("[DEBUG] Se llamó a la macro ActualizarDatosCriticosEnHojas.")
                except Exception as e:
                    print(f"[ERROR] No se pudo ejecutar la macro: {e}")

                # 5) Guardar el archivo final con el nuevo formato
                nombre_nuevo = f"{registro['id']}_Tu Registro Pedagogico - {profesor.upper()} - {nivel}.xlsm"
                ruta_nueva = os.path.join(carpeta_unidad, nombre_nuevo)
                workbook.SaveAs(ruta_nueva, FileFormat=52)  # 52 = xlsm
                workbook.Close(SaveChanges=False)

                archivos_generados.append(ruta_nueva)

                if os.path.exists(ruta_copia_temporal):
                    os.remove(ruta_copia_temporal)
                    print("[DEBUG] Se eliminó la copia temporal.")

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

    except Exception as e:
        print(f"[ERROR] Error general al generar archivos: {e}")
    finally:
        if excel:
            excel.Quit()
            del excel
        gc.collect()

        # Actualizar el estado en la API (solo si hay IDs que se procesaron correctamente)
        if ids_procesados:
            actualizar_estado_api(ids_procesados)
        else:
            print("[DEBUG] No hubo IDs para actualizar en la API.")
