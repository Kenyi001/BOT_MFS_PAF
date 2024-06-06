# extraccion_datos_web.py
import os
import time
import threading
import pandas as pd
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtWidgets import QMessageBox, QHBoxLayout, QWidget, QLabel
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH, URL_REGISTRO_PAF

cancel_event = threading.Event()

def mostrar_ventana_emergente(mensaje, imagen):
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Notificación")

    custom_widget = QWidget()
    layout = QHBoxLayout()

    pixmap = QPixmap(imagen).scaled(64, 64, aspectRatioMode=Qt.KeepAspectRatio)
    image_label = QLabel()
    image_label.setPixmap(pixmap)
    image_label.setAlignment(Qt.AlignCenter)
    layout.addWidget(image_label)

    text_label = QLabel(mensaje)
    text_label.setFont(QFont("Arial", 10))
    text_label.setAlignment(Qt.AlignCenter)
    layout.addWidget(text_label)

    custom_widget.setLayout(layout)

    msgBox.layout().addWidget(custom_widget)
    msgBox.exec_()

def configurar_driver():
    """Configura el WebDriver de Edge para modo headless."""
    options = Options()
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')
    options.add_argument("--headless")  # Ejecutar en modo headless
    options.add_argument("--disable-gpu")  # Desactivar el GPU
    service = Service(WEBDRIVER_PATH)
    return webdriver.Edge(service=service, options=options)

def iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena):
    """Inicia sesión en la página web con el usuario y contraseña proporcionados."""
    driver.get(url_inicio_sesion)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtUsuario")))
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtPassword")))
    campo_usuario = driver.find_element(By.ID, "MainContent_DefaultContent_txtUsuario")
    campo_contrasena = driver.find_element(By.ID, "MainContent_DefaultContent_txtPassword")
    boton_inicio_sesion = driver.find_element(By.ID, "MainContent_DefaultContent_LoginButton")
    campo_usuario.send_keys(usuario)
    campo_contrasena.send_keys(contrasena)
    boton_inicio_sesion.click()

def buscar_y_editar_mef(driver, mef):
    """Busca y edita el MEF especificado en el formulario web."""
    intentos = 1  # Número de reintentos permitidos
    for intento in range(intentos):
        if cancel_event.is_set():
            print("Cancelación detectada. Saliendo de buscar_y_editar_mef.")
            return False
        try:
            driver.get(URL_REGISTRO_PAF)
            WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body")))
            print("Navegación a la URL completada en intento", intento + 1)
            wait = WebDriverWait(driver, 20)
            wait.until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")))
            print("Elemento de entrada del MEF encontrado.")
            input_mef = driver.find_element(By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")
            input_mef.clear()
            input_mef.send_keys(mef)
            print("MEF ingresado.")
            time.sleep(7)
            wait.until(EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnEditar")))
            print("Botón de edición es clicable.")
            boton_editar = driver.find_element(By.ID, "MainContent_DefaultContent_btnEditar")
            boton_editar.click()
            time.sleep(5)
            print("Botón de edición clicado.")
            return True
        except TimeoutException:
            print("Timeout: No se pudo encontrar el campo MEF o el botón de editar.")
            time.sleep(3)
        except (NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"Problema al interactuar con la página: {str(e)}")
            break
        except Exception as e:
            print(f"Error al buscar y editar el MEF: {str(e)}")
            break
    print("No se pudo completar la acción después de varios intentos.")
    return False

def obtener_dato(driver, selector, method):
    """Obtiene datos de un elemento web dependiendo del método especificado."""
    try:
        if method == "select":
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
            elemento = driver.find_element(*selector)
            return elemento.find_element(By.CSS_SELECTOR, "option:checked").text
        elif method == "input":
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
            input_element = driver.find_element(*selector)
            return input_element.get_attribute('value')
        elif method == "textarea":
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
            textarea_element = driver.find_element(*selector)
            return textarea_element.get_attribute('value')
        elif method == "checkbox":
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
            checkbox_element = driver.find_element(*selector)
            is_checked = "dxWeb_edtCheckBoxChecked" in checkbox_element.get_attribute("class")
            return 1 if is_checked else 0
        elif method == "span":
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
            elemento = driver.find_element(*selector)
            return elemento.text
    except Exception as e:
        print(f"Error obteniendo dato {method} para {selector}: {e}")
        return None

def obtener_datos_horarios_atencion(driver, selector):
    """Obtiene los datos de horarios de atención de una tabla web."""
    try:
        tabla = driver.find_element(*selector)
        filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")
        horarios_atencion = []
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            datos_fila = [celda.text for celda in celdas if celda.get_attribute('class') != 'dxgvCommandColumn dxgv dx-ac']
            horarios_atencion.append(datos_fila)
        return horarios_atencion
    except Exception as e:
        print(f"Error obteniendo horarios de atención: {e}")
        return []

def obtener_datos_horarios_atencion_dia(driver, dia, selector):
    """Obtiene los horarios de atención de un día específico de una tabla web."""
    try:
        tabla = driver.find_element(*selector)
        filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if celdas[0].text == dia:
                horarios = celdas[2:6]
                if horarios[2].text:
                    horario_formateado = f"{horarios[0].text} a {horarios[1].text} y {horarios[2].text} a {horarios[3].text}"
                else:
                    horario_formateado = f"{horarios[0].text} a {horarios[3].text}"
                return horario_formateado
        return "00:00 a 00:00"
    except Exception as e:
        print(f"Error obteniendo horario de atención para {dia}: {e}")
        return "00:00 a 00:00"

def obtener_ubicacion_geografica(driver, selector):
    """Obtiene y divide la ubicación geográfica en 'Departamento', 'Municipio' y 'Localidad'."""
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
        elemento = driver.find_element(*selector)
        ubicacion = elemento.text
        partes = ubicacion.split("\\")
        return {
            "Departamento": partes[2] if len(partes) > 2 else None,
            "Municipio": partes[3] if len(partes) > 3 else None,
            "Localidad": partes[-1] if len(partes) > 0 else None
        }
    except Exception as e:
        print(f"Error obteniendo ubicación geográfica: {e}")
        return {
            "Departamento": None,
            "Municipio": None,
            "Localidad": None
        }

def extraer_datos_web(update_progress, url, archivo_excel, usuario, contrasena):
    cancel_event.clear()
    driver = None
    try:
        driver = configurar_driver()
        iniciar_sesion(driver, url, usuario, contrasena)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "LoginName1")))
        print("Inicio de sesión exitoso.")
        df = pd.read_excel(archivo_excel)
        total_rows = len(df)
        datos_extraidos = []
        errores = []

        for index, row in df.iterrows():
            if cancel_event.is_set():
                print("Cancelación detectada. Saliendo del bucle de extracción de datos.")
                break
            mef_puro = row['MEF']
            mef = mef_puro[:3] + "-" + mef_puro[-4:]
            try:
                if not buscar_y_editar_mef(driver, mef):
                    print(f"MEF {mef} no encontrado. Continuando con el siguiente.")
                    errores.append(mef_puro)
                    datos_extraidos.append({"MEF": mef_puro})
                    continue
                datos = {
                    "MEF": mef_puro,
                    "Tipo Sucursal": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ddlTipoSucursal"), "select"),
                    "Nombre Responsable": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF"), "input"),
                    "Identificación": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF"), "input"),
                    "Fecha Apertura": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctrlFechaApertura_I"), "input"),
                    "Nombre": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtNombre"), "input"),
                    "Dirección": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtDireccion"), "textarea"),
                    "Nivel Seguridad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ddlNivelSeguridad"), "select"),
                    "Número de Cámaras": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCamaras"), "input"),
                    "Número de Policias": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtnPolicias"), "input"),
                    "Número de Personal de Seguridad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtnPersonasSeguridad"), "input"),
                    "Teléfono": obtener_dato(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I"), "input"),
                    "Fax": obtener_dato(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I"), "input"),
                    "Horarios Atención": obtener_datos_horarios_atencion(driver, (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Lunes": obtener_datos_horarios_atencion_dia(driver, 'Lunes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Martes": obtener_datos_horarios_atencion_dia(driver, 'Martes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Miercoles": obtener_datos_horarios_atencion_dia(driver, 'Miércoles', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Jueves": obtener_datos_horarios_atencion_dia(driver, 'Jueves', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Viernes": obtener_datos_horarios_atencion_dia(driver, 'Viernes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Sabado": obtener_datos_horarios_atencion_dia(driver, 'Sábado', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Domingo": obtener_datos_horarios_atencion_dia(driver, 'Domingo', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Emitir billeteras y operar cuentas": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D"), "checkbox"),
                    "Ejecutar electrónicamente órdenes de pago y consultas con dispositivos móviles": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D"), "checkbox"),
                    "Operar servicios de pago móvil": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D"), "checkbox"),
                    "Ubicación Geográfica": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), "span"),
                    "Departamento": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), "span"),
                    "Municipio": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), "span"),
                    "Localidad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), "span"),
                    "Coordenada X": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaX"), "input"),
                    "Coordenada Y": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaY"), "input")
                }
                datos.update(obtener_ubicacion_geografica(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica")))
                datos_extraidos.append(datos)
            except Exception as e:
                print(f"Error al extraer datos del MEF {mef}: {str(e)}")
                errores.append(mef_puro)
                datos_extraidos.append({"MEF": mef_puro, "Error": str(e)})

            # Actualizar el progreso
            progress_value = int((index + 1) / total_rows * 100)
            update_progress(progress_value)

        if not cancel_event.is_set():
            df_extraido = pd.DataFrame(datos_extraidos)
            ruta_carpeta = os.path.join(os.getcwd(), 'Reporte data PTM')
            if not os.path.exists(ruta_carpeta):
                os.makedirs(ruta_carpeta)
            formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
            nombre_archivo = f"PTM_extraidos_{formato_fecha_hora}.xlsx"
            nombre_archivo_salida = os.path.join(ruta_carpeta, nombre_archivo)
            df_extraido.to_excel(nombre_archivo_salida, index=False)
            print(f"Datos extraídos guardados en '{nombre_archivo_salida}'.")

    except Exception as e:
        print(f"Ocurrió un error inesperado: {str(e)}")
    finally:
        if driver:
            driver.quit()
        if errores:
            mostrar_ventana_emergente(
                f"Se generaron algunos errores en la extracción de los MEF: {', '.join(errores)}. Se guardaron en un archivo Excel.",
                os.path.join('resources', 'error.webp')
            )
