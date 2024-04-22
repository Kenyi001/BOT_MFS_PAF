# new_PAF_PTM.py
import os
from datetime import datetime

import pandas as pd
import openpyxl
# Importaciones adicionales
from flask import jsonify, Response
from PyQt5.QtWidgets import QFileDialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

def iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena):
    driver.get(url_inicio_sesion)

    # Esperar hasta que los campos de usuario y contraseña estén presentes
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtUsuario")))
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtPassword")))

    # Llenar y enviar el formulario de inicio de sesión
    campo_usuario = driver.find_element(By.ID, "MainContent_DefaultContent_txtUsuario")
    campo_contrasena = driver.find_element(By.ID, "MainContent_DefaultContent_txtPassword")
    boton_inicio_sesion = driver.find_element(By.ID, "MainContent_DefaultContent_LoginButton")

    campo_usuario.send_keys(usuario)
    campo_contrasena.send_keys(contrasena)
    boton_inicio_sesion.click()
def seleccionar_ruta_guardado():
    dialogo = QFileDialog()
    ruta_guardado = dialogo.getSaveFileName(caption="Guardar Archivo de Errores", directory=obtener_ruta_escritorio(), filter="Excel Files (*.xlsx)")
    return ruta_guardado[0]  # Retorna la ruta del archivo seleccionado
def obtener_ruta_escritorio():
    ruta_escritorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    return ruta_escritorio
def agrupar_dias_con_horarios_comunes(dias):
    grupos = {}
    for num_dia, horarios in dias.items():
        tipo_horario, horarios_dia = determinar_tipo_horario(horarios)

        # Saltar los días "Sin Atención"
        if tipo_horario == "-Sin Atención-":
            continue

        clave = (tipo_horario, tuple(horarios_dia))  # Clave única para cada combinación de tipo de horario y horarios

        if clave not in grupos:
            grupos[clave] = []
        grupos[clave].append(num_dia)

    return grupos
def establecer_coordenadas(driver, coordenada_x, coordenada_y, id_coordenada_x, id_coordenada_y):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, id_coordenada_x)))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, id_coordenada_y)))

    driver.execute_script(f"document.getElementById('{id_coordenada_x}').removeAttribute('disabled');")
    driver.execute_script(f"document.getElementById('{id_coordenada_y}').removeAttribute('disabled');")

    driver.execute_script(f"document.getElementById('{id_coordenada_x}').value = '{coordenada_x}';")
    driver.execute_script(f"document.getElementById('{id_coordenada_y}').value = '{coordenada_y}';")
def buscar_localidad(driver, localidad_id, boton_buscar_id, localidad):
    campo_localidad = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, localidad_id))
    )
    campo_localidad.clear()

    # Quitar solo el primer espacio si localidad comienza con un espacio
    localidad_modificada = localidad[1:] if localidad.startswith(' ') else localidad

    campo_localidad.send_keys(localidad_modificada)

    boton_buscar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, boton_buscar_id))
    )
    boton_buscar.click()

def seleccionar_localidad(driver, tabla_id, localidad_buscada, departamento_verificado):
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, tabla_id))
    )

    # Quitar solo el primer espacio si localidad_buscada comienza con un espacio
    localidad_buscada = localidad_buscada[1:] if localidad_buscada.startswith(' ') else localidad_buscada
    # Aplica también la misma lógica a departamento_verificado si crees que podría ser necesario
    departamento_verificado = departamento_verificado[1:] if departamento_verificado.startswith(' ') else departamento_verificado

    tabla = driver.find_element(By.ID, tabla_id)
    filas = tabla.find_elements(By.TAG_NAME, "tr")
    localidad_seleccionada = False  # Inicializa una bandera para seguir si se seleccionó alguna localidad

    for fila in filas[1:]:  # Comenzar desde 1 para saltar la fila del encabezado
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if len(celdas) > 1:
            nombre_localidad = celdas[1].text.strip()
            datos_ubicacion = celdas[2].text.strip().split('\\')
            if len(datos_ubicacion) >= 3:
                departamento = datos_ubicacion[2].strip()  # Elimina espacios adicionales

                if nombre_localidad.lower() == localidad_buscada.strip().lower() and departamento.lower() == departamento_verificado.strip().lower():
                    print(f"Coincidencia encontrada: {nombre_localidad}, {departamento}. Intentando hacer clic...")
                    boton_seleccionar = WebDriverWait(celdas[0], 10).until(
                        EC.element_to_be_clickable((By.TAG_NAME, "a"))
                    )
                    boton_seleccionar.click()
                    localidad_seleccionada = True
                    break

    if not localidad_seleccionada:
        print(f"No se encontró la localidad '{localidad_buscada.strip()}' con departamento '{departamento_verificado.strip()}'. Deteniendo el bot.")


def seleccionar_dias(driver, dias):
    for dia in dias:
        checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"td:nth-child({dia}) > label")))
        checkbox.click()
def determinar_tipo_horario(horarios_dia):
    manana_ini, manana_fin, tarde_ini, tarde_fin = horarios_dia

    if manana_ini == "0:00" and manana_fin == "0:00" and tarde_ini == "" and tarde_fin == "":
        return "-Sin Atención-", horarios_dia
    elif manana_ini == "00:01" and manana_fin == "00:01" and tarde_ini == "" and tarde_fin == "":
        return "24 horas", horarios_dia
    elif tarde_ini == tarde_fin:
        return "Continuo", horarios_dia
    else:
        return "Discontinuo", horarios_dia

def establecer_tipo_horario(driver, tipo_horario_id, tipo_horario_text):
    # Esperar hasta que el elemento del menú desplegable sea clickable
    select_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, tipo_horario_id))
    )
    # Crear un objeto Select y seleccionar por texto visible
    Select(select_element).select_by_visible_text(tipo_horario_text)
def establecer_horario(driver, tipo_horario, horarios, prefijo_id):
    # Seleccionar tipo de horario
    establecer_tipo_horario(driver, f'{prefijo_id}ddlTipoHorario', tipo_horario)

    if tipo_horario in ["24 horas", "-Sin Atención-"]:
        # No se configuran horarios específicos para estos casos
        return

    entrada1, salida1, entrada2, salida2 = horarios

    # Configurar los horarios para el tipo "Continuo"
    if tipo_horario == "Continuo":
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, f"{prefijo_id}timeEntrada1_I")))
        driver.find_element(By.ID, "MainContent_DefaultContent_timeEntrada1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada1_I').value = '{entrada1}';")
        driver.find_element(By.ID, "MainContent_DefaultContent_timeSalida2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida2_I').value = '{salida1}';")

    # Configurar los horarios para el tipo "Discontinuo"
    elif tipo_horario == "Discontinuo":
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, f"{prefijo_id}timeEntrada1_I")))
        driver.find_element(By.ID, "MainContent_DefaultContent_timeEntrada1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada1_I').value = '{entrada1}';")
        driver.find_element(By.ID, "MainContent_DefaultContent_timeSalida1_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida1_I').value = '{salida1}';")
        driver.find_element(By.ID, "MainContent_DefaultContent_timeEntrada2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada2_I').value = '{entrada2}';")
        driver.find_element(By.ID, "MainContent_DefaultContent_timeSalida2_I").click()
        driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida2_I').value = '{salida2}';")

def guardar_horario(driver, tipo_horario, horario_dia, prefijo_id):
    if tipo_horario != "-Sin Atención-":
        establecer_horario(driver, tipo_horario, horario_dia, prefijo_id)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, f"{prefijo_id}btnAdicionarHorario"))
        ).click()

        # Determinar si es necesario realizar acciones adicionales basadas en el horario
        realizar_accion_adicional = False
        if tipo_horario == "Continuo" and horario_dia[1] >= "23:00":
            realizar_accion_adicional = True
        elif tipo_horario == "Discontinuo" and horario_dia[3] >= "23:00":
            realizar_accion_adicional = True

        # Acciones adicionales si se cumple la condición
        if realizar_accion_adicional:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, f"{prefijo_id}ASPxPopupControlMensaje_HCB-1"))
            ).click()

# Función para dividir los horarios en horario de mañana y tarde
def dividir_horarios(horario):
    # Manejar casos nulos o especiales
    if pd.isnull(horario) or horario in ["", "0", "0:00", "00:00", "00:00 a 00:00", "00:00 a 00:00 y 00:00 a 00:00"]:
        return "-Sin Atención-", ["0:00", "0:00", "", ""]

    # Dividir horario en partes (mañana y tarde si es necesario)
    if "y" in horario:
        partes = horario.split(' y ')
        manana = partes[0].split(' a ')
        tarde = partes[1].split(' a ')
        horarios_dia = manana + tarde
    else:
        partes = horario.split(' a ')
        horarios_dia = partes + ["", ""]  # Asegura que siempre hay 4 elementos

    # Asegurarse de que horarios_dia tenga siempre 4 elementos
    while len(horarios_dia) < 4:
        horarios_dia.append("")

    # Determinar el tipo de horario
    manana_ini, manana_fin, tarde_ini, tarde_fin = horarios_dia
    if manana_ini == "00:01" and manana_fin == "00:01" and tarde_ini == "" and tarde_fin == "":
        return "24 horas", horarios_dia
    elif tarde_ini == tarde_fin:
        return "Continuo", horarios_dia
    else:
        return "Discontinuo", horarios_dia
def determinar_horario_comun(*horarios_semana):
    # Supongamos que cada entrada en horarios_semana es una lista de horarios [inicio1, fin1, inicio2, fin2]
    horario_referencia = horarios_semana[0]
    for horario in horarios_semana[1:]:
        if horario != horario_referencia:
            return None  # Si algún horario es diferente, no hay horario común
    return horario_referencia  # Todos los horarios son iguales


def ejecutar_automatizacion_new(url_inicio_sesion, ruta_archivo_excel, usuario, contrasena):
    # Configuración inicial para modo headless
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')  # Si necesitas cargar un perfil de usuario específico
    service = Service(WEBDRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)

    print("Iniciando el navegador Edge en modo headless...")

    # Iniciar sesión en la plataforma
    # url_inicio_sesion = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
    iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena)

    print("Inicio de sesión exitoso.")

    # Aquí continúa el procesamiento de tu archivo Excel y la automatización subsecuente
    try:
        df = pd.read_excel(
            ruta_archivo_excel,
            sheet_name="Data",  # Nombre de la Hoja de Cálculo
            dtype=str  # Forzar todas las columnas a ser tratadas como texto
        )
        print("Lectura de archivo Excel exitosa.")
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        return  # Finaliza la ejecución si hay un error al leer el archivo
    contador_row = 0
    errores = []
    max_retries = 3

    for index, row in df.iterrows():
        try:
            contador_row += 1
            # Accede a los valores de las columnas que necesitas
            numero_MEF = row["mef"]
            tipo_sucursal = "Punto de Atención Corresponsal No Financiero"
            nombre_responsable = row["nombre_corresponsable"]
            identificacion_responsable = row["nro_carnet"]
            nombre_comercio = row["nombre_del_negocio"]
            direccion = row["direccion"]
            nivel_seguridad = "Bajo"
            camaras_seguridad = row["camara_de_seguridad"]
            telefono = row["linea_personal_corresponsable"]
            fax = "0"

            # Coordenadas
            coordenada_x = row["lat_gps"]
            coordenada_y = row["lng_gps"]

            # Servicios
            # servicios_seleccionados = [0, 1, 2]

            # Localidad
            localidad_busqueda = row["Localidad"]
            departamento_verificado = row["departamento"]

            # Mensaje deseado
            mensaje_deseado = "Se guardo correctamente el punto de atención"

            # Navegar a un sitio web de Formulario
            driver.get("https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx")

            # Acciones de clic y tipo
            driver.find_element(By.ID, "MainContent_DefaultContent_btnAdicionar").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNumero").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNumero").send_keys(numero_MEF)
            driver.find_element(By.ID, "MainContent_DefaultContent_ddlTipoSucursal").click()
            # Selección de un elemento en un dropdown (seleccionar por texto visible)
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlTipoSucursal'))
            select.select_by_visible_text(tipo_sucursal)

            # Esperar para visualizar la página
            driver.implicitly_wait(30)  # O ajusta este tiempo según sea necesario

            # Interacción con campos de texto y selección de opciones
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF").send_keys(nombre_responsable)
            driver.find_element(By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF").send_keys(identificacion_responsable)

            # espera que se cargue el elemento Nombre de Comercio
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_txtNombre")))
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNombre").send_keys(nombre_comercio)

            # espera que se cargue el elemento descripcion de direccion
            # Identifica el elemento por su ID
            id_elemento = "MainContent_DefaultContent_txtDireccion"


            # Ejecuta un script de JavaScript para cambiar el valor del elemento
            driver.execute_script(f"document.getElementById('{id_elemento}').value = '{direccion}';")

            # Interactuar con dropdowns
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlNivelSeguridad'))
            select.select_by_visible_text(nivel_seguridad)

            # Interacción con campos de teléfono y fax
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I").send_keys(telefono)
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I").send_keys(fax)

            # Interaccion con el campo de Número de Cámaras de Seguridad
            # Localizar el campo de entrada de cámaras de seguridad por ID
            campo_camaras_seguridad = driver.find_element(By.ID, "MainContent_DefaultContent_txtCamaras")
            campo_camaras_seguridad.clear()
            campo_camaras_seguridad.send_keys(camaras_seguridad)

            # Horarios para cada día
            horario_LUN = dividir_horarios(row["horario_lunes"])
            horario_MAR = dividir_horarios(row["horario_martes"])
            horario_MIE = dividir_horarios(row["horario_miercoles"])
            horario_JUE = dividir_horarios(row["horario_jueves"])
            horario_VIE = dividir_horarios(row["horario_viernes"])
            horario_SAB = dividir_horarios(row["horario_sabado"])
            horario_DOM = dividir_horarios(row["horario_domingo"])

            # Procesar y guardar horarios para cada día individualmente
            dias_horarios = [horario_LUN, horario_MAR, horario_MIE, horario_JUE, horario_VIE, horario_SAB, horario_DOM]


            for num_dia, horarios_dia in enumerate(dias_horarios, start=1):

                tipo_horario, horarios = horarios_dia

                if tipo_horario != "-Sin Atención-":
                    # Seleccionar el día
                    seleccionar_dias(driver, [num_dia])

                    guardar_horario(driver, tipo_horario, horarios, 'MainContent_DefaultContent_')

            # Seleccionar servicios
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D").click()

            # Uso de las funciones
            localidad_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_txtLocalidad"
            boton_buscar_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_btnBuscar"
            tabla_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_gridGeografiaLocalidades"

            buscar_localidad(driver, localidad_id, boton_buscar_id, localidad_busqueda)
            seleccionar_localidad(driver, tabla_id, localidad_busqueda, departamento_verificado)

            # Habilitar y llenar campos de coordenadas
            id_coordenada_x = "MainContent_DefaultContent_txtCoordenadaX"
            id_coordenada_y = "MainContent_DefaultContent_txtCoordenadaY"
            establecer_coordenadas(driver, coordenada_x, coordenada_y, id_coordenada_x, id_coordenada_y)

            # time.sleep(20)

            # Finalizar y enviar el formulario
            driver.find_element(By.ID, "MainContent_DefaultContent_btnGrabar").click()
            # Esperar hasta que el elemento con el mensaje se haga visible
            wait = WebDriverWait(driver, 10)
            mensaje_element = wait.until(
                EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_lblMensaje")))  # Comprobador de Exito

            # Obtener el texto del mensaje
            mensaje_texto = mensaje_element.text

            # Verificar si el mensaje deseado está presente en el texto
            if mensaje_deseado in mensaje_texto:
                print("El proceso se completó correctamente.")
            else:
                response = jsonify({'error': 'An error occurred', 'details': f'Error at MEF number {numero_MEF}: {str(errores)}', 'mef': numero_MEF})
                return Response(response, status=502)
        except Exception as e:
            print(f"Error at MEF number {numero_MEF}: {str(e)}")
            print(f"Exception type: {type(e)}")
            import traceback
            print("Exception traceback:")
            print(traceback.format_exc())
            errores.append(f"Error en la fila {contador_row}: {str(e)}")
            response = jsonify({'error': 'An error occurred while processing the files', 'details': errores, 'mef': numero_MEF})
            return Response(response, status=501)

    response = jsonify({'message': 'All files processed successfully'})
    return Response(response, status=200)
    driver.quit()