import os
import time
import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from datetime import datetime
from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH, URL_REGISTRO_PAF

# Obtener la ruta del directorio actual
ruta_actual = os.getcwd()

def mostrar_ventana_emergente():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Critical)  # Establecer el icono a rojo
    msgBox.setWindowTitle("Notificacion de Error")
    msgBox.setText("Se generaron algunos errores en la Extraccion de los MEF y se guardaron en un archivo Excel.")
    msgBox.exec()
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

    # Esperar hasta que los campos de usuario y contraseña estén presentes
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtUsuario")))
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtPassword")))

    # Llenar y enviar el formulario de inicio de sesión
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
        try:
            driver.get(URL_REGISTRO_PAF)
            WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body")))
            #time.sleep(5)
            print("Navegación a la URL completada en intento", intento + 1)

            wait = WebDriverWait(driver, 20)
            wait.until(
                EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")))
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
            return
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
def obtener_dato_select(driver, selector):
    # Espera hasta que el elemento esté presente
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
    # Encuentra el elemento por su selector
    elemento = driver.find_element(*selector)
    # Obtiene el texto del elemento seleccionado
    return elemento.find_element(By.CSS_SELECTOR, "option:checked").text
def obtener_dato_input(driver, selector):
    # Encuentra el elemento input por su selector
    input_element = driver.find_element(*selector)
    # Obtiene el valor del atributo 'value'
    return input_element.get_attribute('value')
def obtener_dato_textarea(driver, selector):
    # Encuentra el elemento textarea por su selector
    textarea_element = driver.find_element(*selector)
    # Obtiene el texto contenido en el textarea
    return textarea_element.get_attribute('value')
def obtener_dato_checkbox(driver, selector):
    # Encuentra el elemento span por su selector
    checkbox_element = driver.find_element(*selector)
    # Verifica la clase del elemento para determinar si está "presionado"
    is_checked = "dxWeb_edtCheckBoxChecked" in checkbox_element.get_attribute("class")
    # Retorna 1 si está "presionado", 0 si no lo está
    return 1 if is_checked else 0
def obtener_dato_span(driver, selector):
    # Encuentra el elemento por su selector
    elemento = driver.find_element(*selector)
    # Obtiene el texto del elemento
    return elemento.text
def obtener_datos_horarios_atencion(driver, selector):
    # Encuentra la tabla por su ID
    tabla = driver.find_element(*selector)
    # Encuentra todas las filas de datos en la tabla
    filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")

    horarios_atencion = []  # Lista para almacenar los horarios

    # Recorre cada fila para extraer los datos
    for fila in filas:
        # Encuentra todas las celdas de la fila
        celdas = fila.find_elements(By.TAG_NAME, "td")
        # Extrae el texto de cada celda y lo añade a una lista
        datos_fila = [celda.text for celda in celdas if celda.get_attribute('class') != 'dxgvCommandColumn dxgv dx-ac']
        # Añade la lista de datos de la fila a la lista de horarios
        horarios_atencion.append(datos_fila)

    return horarios_atencion
def obtener_datos_horarios_atencion_dia(driver, dia, selector):
    # Encuentra la tabla por su ID
    tabla = driver.find_element(*selector)
    # Encuentra todas las filas de datos en la tabla
    filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")

    # Recorre cada fila para extraer los datos
    for fila in filas:
        # Encuentra todas las celdas de la fila
        celdas = fila.find_elements(By.TAG_NAME, "td")
        # Encuentra la celda correspondiente al día
        if celdas[0].text == dia:
            # Extrae los horarios de la celda
            horarios = celdas[2:6]  # Las celdas 2 a 5 contienen los horarios
            # Formatea los horarios en el formato deseado
            if horarios[2].text:  # Si hay un segundo horario
                horario_formateado = f"{horarios[0].text} a {horarios[1].text} y {horarios[2].text} a {horarios[3].text}"
            else:  # Si solo hay un horario
                horario_formateado = f"{horarios[0].text} a {horarios[1].text}"
            return horario_formateado

    # Si el día no está en la tabla, devuelve "00:00 a 00:00"
    return "00:00 a 00:00"
def obtener_dato_span_Dep(driver, selector):
    # Encuentra el elemento por su selector
    elemento = driver.find_element(*selector)
    # Obtiene el texto del elemento
    ubicacion = elemento.text
    # Divide la ubicación en sus componentes
    partes = ubicacion.split("\\")
    # Devuelve el segundo elemento (el departamento)
    return partes[2]
def obtener_dato_span_Mun(driver, selector):
    # Encuentra el elemento por su selector
    elemento = driver.find_element(*selector)
    # Obtiene el texto del elemento
    ubicacion = elemento.text
    # Divide la ubicación en sus componentes
    partes = ubicacion.split("\\")
    # Devuelve el segundo elemento (el departamento)
    return partes[3]

def obtener_dato_span_Loc(driver, selector):
    # Encuentra el elemento por su selector
    elemento = driver.find_element(*selector)
    # Obtiene el texto del elemento
    ubicacion = elemento.text
    # Divide la ubicación en sus componentes
    partes = ubicacion.split("\\")
    # Devuelve el último elemento (la localidad)
    return partes[-1]
def extraer_datos_web(url, archivo_excel, usuario, contrasena):
    try:
        driver = configurar_driver()
        iniciar_sesion(driver, url, usuario, contrasena)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "LoginName1")))
        print("Inicio de sesión exitoso.")
        # Leer los datos del archivo Excel
        df = pd.read_excel(archivo_excel)
        datos_extraidos = []
        error = []

        for _, row in df.iterrows():
            mef_puro = row['MEF']
            mef = mef_puro[:3] + "-" + mef_puro[-4:]
            try:
                buscar_y_editar_mef(driver, mef)

                # Extracción de datos (IDs proporcionados)
                datos = {
                    "MEF": mef_puro,
                    "Tipo Sucursal": obtener_dato_select(driver, (By.ID, "MainContent_DefaultContent_ddlTipoSucursal")),
                    "Nombre Responsable": obtener_dato_input(driver,
                                                      (By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF")),
                    "Identificación": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF")),
                    "Fecha Apertura": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_ctrlFechaApertura_I")),
                    "Nombre": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtNombre")),
                    "Dirección": obtener_dato_textarea(driver, (By.ID, "MainContent_DefaultContent_txtDireccion")),
                    "Nivel Seguridad": obtener_dato_select(driver, (By.ID, "MainContent_DefaultContent_ddlNivelSeguridad")),
                    "Número de Cámaras": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtCamaras")),
                    "Número de Policias": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtnPolicias")),
                    "Número de Personal de Seguridad": obtener_dato_input(driver,
                                                      (By.ID, "MainContent_DefaultContent_txtnPersonasSeguridad")),
                    "Teléfono": obtener_dato_input(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I")),
                    "Fax": obtener_dato_input(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I")),
                    "Horarios Atención": obtener_datos_horarios_atencion(driver,
                                                      (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Lunes": obtener_datos_horarios_atencion_dia(driver, 'Lunes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Martes": obtener_datos_horarios_atencion_dia(driver, 'Martes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Miercoles": obtener_datos_horarios_atencion_dia(driver, 'Miércoles', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Jueves": obtener_datos_horarios_atencion_dia(driver, 'Jueves', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Viernes": obtener_datos_horarios_atencion_dia(driver, 'Viernes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Sabado": obtener_datos_horarios_atencion_dia(driver, 'Sabado', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Domingo": obtener_datos_horarios_atencion_dia(driver, 'Domingo', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Emitir billeteras y operar cuentas": obtener_dato_checkbox(driver,
                                               (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D")),
                    "Ejecutar electrónicamente órdenes de pago y consultas con dispositivos móviles": obtener_dato_checkbox(driver,
                                               (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D")),
                    "Operar servicios de pago móvil": obtener_dato_checkbox(driver,
                                               (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D")),
                    "Ubicación Geográfica": obtener_dato_span(driver, (
                    By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica")),
                    "Depatamento": obtener_dato_span_Dep(driver, (
                    By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica")),
                    "Municipio": obtener_dato_span_Mun(driver, (
                    By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica")),
                    "Localidad": obtener_dato_span_Loc(driver, (
                    By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica")),
                    "Coordenada X": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaX")),
                    "Coordenada Y": obtener_dato_input(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaY"))
                }

                datos_extraidos.append(datos)
            except Exception as e:
                print(f"Error al extraer datos del MEF {mef}: {str(e)}")
                datos_extraidos.append({
                    "MEF": mef_puro,
                    "Error": str(e)
                })
                error = 1

        df_extraido = pd.DataFrame(datos_extraidos)

        # Crear carpeta para guardar el archivo si no existe
        ruta_carpeta = os.path.join(ruta_actual, 'Reporte data PTM')
        if not os.path.exists(ruta_carpeta):
            os.makedirs(ruta_carpeta)

        # Generar el nombre del archivo con formato de fecha y hora
        formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
        nombre_archivo = f"PTM_extraidos_{formato_fecha_hora}.xlsx"
        nombre_archivo_salida = os.path.join(ruta_carpeta, nombre_archivo)

        # Guardar los datos extraídos en un nuevo archivo Excel
        df_extraido.to_excel(nombre_archivo_salida, index=False)

        print(f"Datos extraídos guardados en '{nombre_archivo_salida}'.")



    except Exception as e:
        print(f"Ocurrió un error inesperado: {str(e)}")
    finally:
        driver.quit()

        if error:
            mostrar_ventana_emergente()