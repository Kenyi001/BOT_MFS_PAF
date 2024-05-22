import os
import time
import pandas as pd
import json
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from datetime import datetime
from config import URL_REGISTRO_PAF


def load_config():
    """Carga la configuración desde un archivo JSON."""
    default_config = {"WEBDRIVER_PATH": "", "EDGE_BINARY_PATH": "", "USER_PROFILE_PATH": ""}
    config_path = 'config.json'

    if not os.path.exists(config_path):
        with open(config_path, 'w') as config_file:
            json.dump(default_config, config_file)

    with open(config_path, 'r') as config_file:
        return json.load(config_file)


config = load_config()
WEBDRIVER_PATH = config.get('WEBDRIVER_PATH', '')
EDGE_BINARY_PATH = config.get('EDGE_BINARY_PATH', '')
USER_PROFILE_PATH = config.get('USER_PROFILE_PATH', '')


def mostrar_ventana_emergente():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setWindowTitle("Notificación de Error")
    msgBox.setText("Se generaron algunos errores en la extracción de los MEF y se guardaron en un archivo Excel.")
    msgBox.exec()


def configurar_driver():
    """Configura el WebDriver de Edge para modo headless."""
    options = Options()
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    service = Service(WEBDRIVER_PATH)
    return webdriver.Edge(service=service, options=options)


def iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena):
    """Inicia sesión en la página web con el usuario y contraseña proporcionados."""
    driver.get(url_inicio_sesion)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtUsuario")))
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtPassword")))

    driver.find_element(By.ID, "MainContent_DefaultContent_txtUsuario").send_keys(usuario)
    driver.find_element(By.ID, "MainContent_DefaultContent_txtPassword").send_keys(contrasena)
    driver.find_element(By.ID, "MainContent_DefaultContent_LoginButton").click()


def buscar_y_editar_mef(driver, mef):
    """Busca y edita el MEF especificado en el formulario web."""
    for intento in range(2):
        try:
            driver.get(URL_REGISTRO_PAF)
            WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "body")))

            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")))
            input_mef = driver.find_element(By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")
            input_mef.clear()
            input_mef.send_keys(mef)
            time.sleep(7)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnEditar"))).click()
            time.sleep(5)
            return
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"Intento {intento + 1} fallido: {str(e)}")
            time.sleep(3)
        except Exception as e:
            print(f"Error al buscar y editar el MEF: {str(e)}")
            break
    print("No se pudo completar la acción después de varios intentos.")


def obtener_dato(driver, selector, attr='value'):
    """Obtiene el dato de un elemento web."""
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(selector))
    elemento = driver.find_element(*selector)
    if attr == 'text':
        return elemento.text
    elif attr == 'checkbox':
        return 1 if "dxWeb_edtCheckBoxChecked" in elemento.get_attribute("class") else 0
    elif attr == 'select':
        return elemento.find_element(By.CSS_SELECTOR, "option:checked").text
    return elemento.get_attribute('value')


def obtener_datos_horarios_atencion(driver, selector):
    """Obtiene los horarios de atención de una tabla web."""
    tabla = driver.find_element(*selector)
    filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")
    horarios_atencion = []
    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")
        datos_fila = [celda.text for celda in celdas if celda.get_attribute('class') != 'dxgvCommandColumn dxgv dx-ac']
        horarios_atencion.append(datos_fila)
    return horarios_atencion


def obtener_datos_horarios_atencion_dia(driver, dia, selector):
    """Obtiene los horarios de atención de un día específico."""
    tabla = driver.find_element(*selector)
    filas = tabla.find_elements(By.CSS_SELECTOR, "tr.dxgvDataRow")
    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if celdas[0].text == dia:
            horarios = celdas[2:6]
            return f"{horarios[0].text} a {horarios[1].text} y {horarios[2].text} a {horarios[3].text}" if horarios[2].text else f"{horarios[0].text} a {horarios[3].text}"
    return "00:00 a 00:00"


def extraer_datos_web(url, archivo_excel, usuario, contrasena):
    try:
        driver = configurar_driver()
        iniciar_sesion(driver, url, usuario, contrasena)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "LoginName1")))
        print("Inicio de sesión exitoso.")

        df = pd.read_excel(archivo_excel)
        datos_extraidos = []
        error_ocurrido = False

        for _, row in df.iterrows():
            mef_puro = row['MEF']
            mef = mef_puro[:3] + "-" + mef_puro[-4:]
            try:
                buscar_y_editar_mef(driver, mef)

                datos = {
                    "MEF": mef_puro,
                    "Tipo Sucursal": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ddlTipoSucursal"), 'select'),
                    "Nombre Responsable": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF")),
                    "Identificación": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF")),
                    "Fecha Apertura": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctrlFechaApertura_I")),
                    "Nombre": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtNombre")),
                    "Dirección": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtDireccion"), 'text'),
                    "Nivel Seguridad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ddlNivelSeguridad"), 'select'),
                    "Número de Cámaras": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCamaras")),
                    "Número de Policias": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtnPolicias")),
                    "Número de Personal de Seguridad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtnPersonasSeguridad")),
                    "Teléfono": obtener_dato(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I")),
                    "Fax": obtener_dato(driver, (By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I")),
                    "Horarios Atención": obtener_datos_horarios_atencion(driver, (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Lunes": obtener_datos_horarios_atencion_dia(driver, 'Lunes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Martes": obtener_datos_horarios_atencion_dia(driver, 'Martes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Miercoles": obtener_datos_horarios_atencion_dia(driver, 'Miércoles', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Jueves": obtener_datos_horarios_atencion_dia(driver, 'Jueves', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Viernes": obtener_datos_horarios_atencion_dia(driver, 'Viernes', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Sabado": obtener_datos_horarios_atencion_dia(driver, 'Sabado', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Domingo": obtener_datos_horarios_atencion_dia(driver, 'Domingo', (By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable")),
                    "Emitir billeteras y operar cuentas": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D"), 'checkbox'),
                    "Ejecutar electrónicamente órdenes de pago y consultas con dispositivos móviles": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D"), 'checkbox'),
                    "Operar servicios de pago móvil": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D"), 'checkbox'),
                    "Ubicación Geográfica": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), 'text'),
                    "Depatamento": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), 'text').split("\\")[2],
                    "Municipio": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), 'text').split("\\")[3],
                    "Localidad": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_lblUbicacionGeografica"), 'text').split("\\")[-1],
                    "Coordenada X": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaX")),
                    "Coordenada Y": obtener_dato(driver, (By.ID, "MainContent_DefaultContent_txtCoordenadaY"))
                }

                datos_extraidos.append(datos)
            except Exception as e:
                print(f"Error al extraer datos del MEF {mef}: {str(e)}")
                datos_extraidos.append({
                    "MEF": mef_puro,
                    "Error": str(e)
                })
                error_ocurrido = True

        df_extraido = pd.DataFrame(datos_extraidos)

        ruta_carpeta = os.path.join(os.getcwd(), 'Reporte data PTM')
        os.makedirs(ruta_carpeta, exist_ok=True)

        formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
        nombre_archivo_salida = os.path.join(ruta_carpeta, f"PTM_extraidos_{formato_fecha_hora}.xlsx")

        df_extraido.to_excel(nombre_archivo_salida, index=False)

        print(f"Datos extraídos guardados en '{nombre_archivo_salida}'.")

    except Exception as e:
        print(f"Ocurrió un error inesperado: {str(e)}")
    finally:
        driver.quit()
        if error_ocurrido:
            mostrar_ventana_emergente()
