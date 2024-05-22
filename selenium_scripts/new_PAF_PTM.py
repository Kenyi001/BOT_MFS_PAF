# new_PAF_PTM.py

import os
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import json

def load_config():
    # Valores predeterminados para la configuración
    default_config = {"WEBDRIVER_PATH": "", "EDGE_BINARY_PATH": "", "USER_PROFILE_PATH": ""}

    # Si el archivo config.json no existe, lo creamos con los valores predeterminados
    if not os.path.exists('config.json'):
        with open('config.json', 'w') as config_file:
            json.dump(default_config, config_file)

    # Leemos el archivo config.json (que ahora sabemos que existe)
    with open('config.json', 'r') as config_file:
        config = json.load(config_file)

    return config

config = load_config()
WEBDRIVER_PATH = config.get('WEBDRIVER_PATH', '')
EDGE_BINARY_PATH = config.get('EDGE_BINARY_PATH', '')
USER_PROFILE_PATH = config.get('USER_PROFILE_PATH', '')
ruta_actual = os.getcwd()


def mostrar_ventana_emergente():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Critical)  # Establecer el icono a rojo
    msgBox.setWindowTitle("Notificacion de Error")
    msgBox.setText("Se generaron algunos errores en la Extraccion de los MEF y se guardaron en un archivo Excel.")
    msgBox.show()


def esperar_y_encontrar_elemento(driver, by, identificador, tiempo_espera=20):
    """Espera hasta que un elemento sea localizable y luego lo retorna."""
    return WebDriverWait(driver, tiempo_espera).until(
        EC.presence_of_element_located((by, identificador))
    )


def iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena):
    """Inicia sesión en un sitio web proporcionado utilizando credenciales de usuario."""
    driver.get(url_inicio_sesion)
    campo_usuario = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_txtUsuario")
    campo_contrasena = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_txtPassword")
    boton_inicio_sesion = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_LoginButton")
    campo_usuario.send_keys(usuario)
    campo_contrasena.send_keys(contrasena)
    boton_inicio_sesion.click()


def establecer_coordenadas(driver, coordenada_x, coordenada_y, id_coordenada_x, id_coordenada_y):
    """Habilita y establece las coordenadas X e Y en los campos especificados mediante ID."""
    esperar_y_encontrar_elemento(driver, By.ID, id_coordenada_x)
    esperar_y_encontrar_elemento(driver, By.ID, id_coordenada_y)
    script = (
        f"document.getElementById('{id_coordenada_x}').removeAttribute('disabled');"
        f"document.getElementById('{id_coordenada_y}').removeAttribute('disabled');"
        f"document.getElementById('{id_coordenada_x}').value = '{coordenada_x}';"
        f"document.getElementById('{id_coordenada_y}').value = '{coordenada_y}';"
    )
    driver.execute_script(script)


def buscar_localidad(driver, localidad_id, boton_buscar_id, localidad):
    campo_localidad = esperar_y_encontrar_elemento(driver, By.ID, localidad_id)
    campo_localidad.clear()
    localidad_modificada = localidad[1:] if localidad.startswith(' ') else localidad
    campo_localidad.send_keys(localidad_modificada)
    boton_buscar = esperar_y_encontrar_elemento(driver, By.ID, boton_buscar_id)
    boton_buscar.click()


def seleccionar_localidad(driver, tabla_id, localidad_buscada, departamento_verificado):
    """Selecciona una localidad y departamento en una tabla dada, basado en los valores buscados."""
    tabla = esperar_y_encontrar_elemento(driver, By.ID, tabla_id)
    localidad_buscada = localidad_buscada.strip().lower()
    departamento_verificado = departamento_verificado.strip().lower()
    filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]
    localidad_seleccionada = False

    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if len(celdas) > 1:
            nombre_localidad = celdas[1].text.strip().lower()
            datos_ubicacion = celdas[2].text.strip().split('\\')
            if len(datos_ubicacion) >= 3:
                departamento = datos_ubicacion[2].strip().lower()
                if nombre_localidad == localidad_buscada and departamento == departamento_verificado:
                    print(f"Coincidencia encontrada: {nombre_localidad.title()}, {departamento.title()}. Intentando hacer clic...")
                    boton_seleccionar = esperar_y_encontrar_elemento(celdas[0], By.TAG_NAME, "a", 10)
                    boton_seleccionar.click()
                    localidad_seleccionada = True
                    break

    if not localidad_seleccionada:
        print(f"No se encontró la localidad '{localidad_buscada.title()}' con departamento '{departamento_verificado.title()}'. Deteniendo el bot.")


def seleccionar_dias(driver, dias):
    for dia in dias:
        checkbox = esperar_y_encontrar_elemento(driver, By.CSS_SELECTOR, f"td:nth-child({dia}) > label")
        checkbox.click()


def establecer_tipo_horario(driver, tipo_horario_id, tipo_horario_text):
    select_element = esperar_y_encontrar_elemento(driver, By.ID, tipo_horario_id)
    Select(select_element).select_by_visible_text(tipo_horario_text)


def establecer_horario(driver, tipo_horario, horarios, prefijo_id):
    establecer_tipo_horario(driver, f'{prefijo_id}ddlTipoHorario', tipo_horario)
    if tipo_horario in ["24 horas", "-Sin Atención-"]:
        return

    entrada1, salida1, entrada2, salida2 = horarios
    if tipo_horario == "Continuo":
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeEntrada1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada1_I').value = '{entrada1}';")
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeSalida2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida2_I').value = '{salida1}';")
    elif tipo_horario == "Discontinuo":
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeEntrada1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada1_I').value = '{entrada1}';")
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeSalida1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida1_I').value = '{salida1}';")
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeEntrada2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada2_I').value = '{entrada2}';")
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}timeSalida2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida2_I').value = '{salida2}';")


def guardar_horario(driver, tipo_horario, horario_dia, prefijo_id):
    if tipo_horario != "-Sin Atención-":
        establecer_horario(driver, tipo_horario, horario_dia, prefijo_id)
        esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}btnAdicionarHorario").click()
        realizar_accion_adicional = False
        if tipo_horario == "Continuo" and horario_dia[1] >= "23:00":
            realizar_accion_adicional = True
        elif tipo_horario == "Discontinuo" and horario_dia[3] >= "23:00":
            realizar_accion_adicional = True
        if realizar_accion_adicional:
            esperar_y_encontrar_elemento(driver, By.ID, f"{prefijo_id}ASPxPopupControlMensaje_HCB-1").click()


def dividir_horarios(horario):
    if pd.isnull(horario) or horario in ["", "0", "0:00", "00:00", "00:00 a 00:00", "00:00 a 00:00 y 00:00 a 00:00"]:
        return "-Sin Atención-", ["0:00", "0:00", "", ""]
    if "y" in horario:
        partes = horario.split(' y ')
        manana = partes[0].split(' a ')
        tarde = partes[1].split(' a ')
        horarios_dia = manana + tarde
    else:
        partes = horario.split(' a ')
        horarios_dia = partes + ["", ""]
    while len(horarios_dia) < 4:
        horarios_dia.append("")
    manana_ini, manana_fin, tarde_ini, tarde_fin = horarios_dia
    if manana_ini == "00:01" and manana_fin == "00:01" and tarde_ini == "" and tarde_fin == "":
        return "24 horas", horarios_dia
    elif tarde_ini == tarde_fin:
        return "Continuo", horarios_dia
    else:
        return "Discontinuo", horarios_dia


def guardar_formulario_y_verificar_mensaje(driver, boton_guardar_id, mensaje_id, mensaje_deseado, contador_row, numero_MEF, tiempo_espera=10):
    """Guarda el formulario y verifica si el mensaje deseado está presente."""
    try:
        driver.find_element(By.ID, boton_guardar_id).click()
        mensaje_element = WebDriverWait(driver, tiempo_espera).until(
            EC.visibility_of_element_located((By.ID, mensaje_id))
        )
        mensaje_texto = mensaje_element.text
        if mensaje_deseado in mensaje_texto:
            print(f"El mensaje deseado está presente: {mensaje_texto} - Fila: {contador_row} - MEF: {numero_MEF}")
            return True
        else:
            print(f"El mensaje deseado no está presente: {mensaje_texto} - Fila: {contador_row} - MEF: {numero_MEF}")
            return False
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error al guardar el formulario o al verificar el mensaje: {e} - Fila: {contador_row} - MEF: {numero_MEF}")
        return False


def ejecutar_automatizacion_new(url_inicio_sesion, ruta_archivo_excel, nombre_hoja, usuario, contrasena):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')
    service = Service(WEBDRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)

    print("Iniciando el navegador Edge en modo headless...")
    iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "LoginName1")))
    print("Inicio de sesión exitoso.")

    try:
        df = pd.read_excel(ruta_archivo_excel, sheet_name=nombre_hoja, dtype=str)
        datos_extraidos = []
        print("Lectura de archivo Excel exitosa.")
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        return

    contador_row = 0
    ultima_accion = ""

    for index, row in df.iterrows():
        contador_row += 1
        numero_MEF = row["mef"]
        try:
            tipo_sucursal = "Punto de Atención Corresponsal No Financiero"
            nombre_responsable = row["nombre_corresponsable"]
            identificacion_responsable = row["nro_carnet"]
            nombre_comercio = row["nombre_del_negocio"]
            direccion = row["direccion"]
            nivel_seguridad = "Bajo"
            camaras_seguridad = row["camara_de_seguridad"]
            telefono = row["linea_personal_corresponsable"]
            fax = "0"
            coordenada_x = row["lat_gps"]
            coordenada_y = row["lng_gps"]
            localidad_busqueda = row["Localidad"]
            departamento_verificado = row["departamento"]
            mensaje_deseado = "Se guardo correctamente el punto de atención"

            driver.get("https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx")
            driver.find_element(By.ID, "MainContent_DefaultContent_btnAdicionar").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNumero").send_keys(numero_MEF)
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlTipoSucursal'))
            select.select_by_visible_text(tipo_sucursal)
            driver.implicitly_wait(30)
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF").send_keys(nombre_responsable)
            driver.find_element(By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF").send_keys(identificacion_responsable)
            esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_txtNombre").send_keys(nombre_comercio)
            driver.execute_script(f"document.getElementById('MainContent_DefaultContent_txtDireccion').value = '{direccion}';")
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlNivelSeguridad'))
            select.select_by_visible_text(nivel_seguridad)
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I").send_keys(telefono)
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I").send_keys(fax)
            campo_camaras_seguridad = driver.find_element(By.ID, "MainContent_DefaultContent_txtCamaras")
            campo_camaras_seguridad.clear()
            campo_camaras_seguridad.send_keys(camaras_seguridad)
            dias_horarios = [
                dividir_horarios(row["horario_lunes"]),
                dividir_horarios(row["horario_martes"]),
                dividir_horarios(row["horario_miercoles"]),
                dividir_horarios(row["horario_jueves"]),
                dividir_horarios(row["horario_viernes"]),
                dividir_horarios(row["horario_sabado"]),
                dividir_horarios(row["horario_domingo"])
            ]

            for num_dia, horarios_dia in enumerate(dias_horarios, start=1):
                tipo_horario, horarios = horarios_dia
                if tipo_horario != "-Sin Atención-":
                    seleccionar_dias(driver, [num_dia])
                    guardar_horario(driver, tipo_horario, horarios, 'MainContent_DefaultContent_')

            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D").click()

            buscar_localidad(driver, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_txtLocalidad",
                             "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_btnBuscar", localidad_busqueda)
            seleccionar_localidad(driver, "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_gridGeografiaLocalidades",
                                  localidad_busqueda, departamento_verificado)

            establecer_coordenadas(driver, coordenada_x, coordenada_y,
                                   "MainContent_DefaultContent_txtCoordenadaX", "MainContent_DefaultContent_txtCoordenadaY")

            resultado_guardado = guardar_formulario_y_verificar_mensaje(driver,
                                                                        "MainContent_DefaultContent_btnGrabar",
                                                                        "MainContent_DefaultContent_lblMensaje",
                                                                        mensaje_deseado,
                                                                        contador_row,
                                                                        numero_MEF)
            if resultado_guardado:
                print("Registro guardado exitosamente.")
            else:
                print("Fallo al guardar el registro.")
        except Exception as e:
            print(f"Error al extraer datos del MEF {numero_MEF}: {str(e)}")
            datos_extraidos.append({
                "MEF": numero_MEF,
                "Error": str(e),
                "Ultima Accion": ultima_accion
            })

    driver.quit()

    if datos_extraidos:
        ruta_carpeta = os.path.join(ruta_actual, 'Errores de Carga')
        if not os.path.exists(ruta_carpeta):
            os.makedirs(ruta_carpeta)
        formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
        nombre_archivo = f"PTM_Error_Carga_{formato_fecha_hora}.xlsx"
        nombre_archivo_salida = os.path.join(ruta_carpeta, nombre_archivo)
        df_extraido = pd.DataFrame(datos_extraidos)
        df_extraido.to_excel(nombre_archivo_salida, index=False)
        mostrar_ventana_emergente()
