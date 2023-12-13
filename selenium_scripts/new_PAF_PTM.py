# new_PAF_PTM.py
import os
from datetime import datetime

import pandas as pd
import openpyxl
# Importaciones adicionales
from PyQt5.QtWidgets import QFileDialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH


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
    campo_localidad.send_keys(localidad)

    boton_buscar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, boton_buscar_id))
    )
    boton_buscar.click()
def seleccionar_localidad(driver, tabla_id, localidad_buscada):
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, tabla_id))
    )

    tabla = driver.find_element(By.ID, tabla_id)
    filas = tabla.find_elements(By.TAG_NAME, "tr")

    for fila in filas[1:]:  # Comenzar desde 1 para saltar la fila del encabezado
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if len(celdas) > 1:
            nombre_localidad = celdas[1].text.strip()  # Segunda celda contiene el nombre de la localidad

            # Verifica si el nombre de la localidad coincide exactamente
            if nombre_localidad.lower() == localidad_buscada.lower():
                celdas[0].find_element(By.TAG_NAME, "a").click()  # Hacer clic en el enlace "Seleccionar" en la primera celda
                return
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
def ejecutar_automatizacion_new(ruta_archivo_excel):
    # Inicializa la lista de registros con errores
    registros_con_errores = []
    # Ruta al WebDriver de Edge Dev
    webdriver_path = WEBDRIVER_PATH

    # Ruta al ejecutable de Edge Dev
    edge_path = EDGE_BINARY_PATH

    # Ruta al perfil de usuario específico de Edge Dev (ajusta esto según sea necesario)
    perfil_usuario = USER_PROFILE_PATH

    # Configuración de las opciones para apuntar a Edge Dev y al perfil específico
    options = Options()
    options.binary_location = edge_path
    options.add_argument(f'user-data-dir={perfil_usuario}')

    # Crear el servicio con la ruta del WebDriver
    service = Service(webdriver_path)

    # Iniciar Edge Dev con las opciones configuradas
    print("Iniciando el navegador Edge Dev...")
    driver = webdriver.Edge(service=service, options=options)
    print("Navegador Edge Dev iniciado correctamente.")

    # Lee los datos desde el archivo Excel y forza todas las columnas a ser tratadas como texto
    try:
        df = pd.read_excel(
            ruta_archivo_excel,
            sheet_name="Cargas_de_Alta_BOT",
            dtype=str  # Forza todas las columnas a ser tratadas como texto
        )
        print("Lectura de archivo Excel exitosa.")
    except Exception as e:
        print(f"se abrio correctamente el excel: {e}")

    # URL de la página de inicio de sesión
    url_inicio_sesion = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
    # Navegar a la página de inicio de sesión
    driver.get(url_inicio_sesion)
    #time.sleep(20)

    # Esperar hasta que la URL cambie a la de la página de inicio
    nueva_url = "https://appweb.asfi.gob.bo/RMI/Default.aspx"  # URL después del inicio de sesión
    WebDriverWait(driver, 240).until(EC.url_to_be(nueva_url))

    hora_iniciacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Inicio de sesión exitoso, URL actualizada a las {hora_iniciacion}.")

    # Contador de filas y numero de procesos
    contador_row = 0

    # Inicio de Bucle
    # Itera a través de las filas del DataFrame
    for index, row in df.iterrows():
        try:
            contador_row += 1
            # Accede a los valores de las columnas que necesitas
            numero_MEF = row["MEF"]
            tipo_sucursal = "Punto de Atención Corresponsal No Financiero"
            nombre_responsable = row["Nombre de Corresponsal"]
            identificacion_responsable = row["CI"]
            nombre_comercio = row["Nombre del negocio"]
            direccion = row["Direccion"]
            nivel_seguridad = "Bajo"
            telefono = row["Linea Personal PTM"]
            fax = "0"

            # Coordenadas
            coordenada_x = row["LATITUD_X"]
            coordenada_y = row["LONGITUD_Y"]

            # Servicios
            # servicios_seleccionados = [0, 1, 2]

            # Localidad
            localidad_busqueda = row["Localidad"]

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

            # Horarios para cada día
            horario_LUN = dividir_horarios(row["HORARIO_LUNES"])
            horario_MAR = dividir_horarios(row["HORARIO_MARTES"])
            horario_MIE = dividir_horarios(row["HORARIO_MIERCOLES"])
            horario_JUE = dividir_horarios(row["HORARIO_JUEVES"])
            horario_VIE = dividir_horarios(row["HORARIO_VIERNES"])
            horario_SAB = dividir_horarios(row["HORARIO_SABADO"])
            horario_DOM = dividir_horarios(row["HORARIO_DOMINGO"])

            # Procesar y guardar horarios para cada día individualmente
            dias_horarios = [horario_LUN, horario_MAR, horario_MIE, horario_JUE, horario_VIE, horario_SAB, horario_DOM]

            for num_dia, horarios_dia in enumerate(dias_horarios, start=1):

                tipo_horario, horarios = horarios_dia

                if tipo_horario != "-Sin Atención-":
                    # Seleccionar el día
                    seleccionar_dias(driver, [num_dia])

                    # Establecer y guardar el horario
                    # if tipo_horario not in ["-Sin Atención-", "24 horas"]:
                    #     establecer_horario(driver, tipo_horario, horarios, 'MainContent_DefaultContent_')
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
            seleccionar_localidad(driver, tabla_id, localidad_busqueda)

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
                print("El mensaje deseado está presente:", mensaje_texto, " - ", contador_row, "-",numero_MEF)
            else:
                print("El mensaje deseado no está presente:", mensaje_texto, " - ", contador_row, "-",numero_MEF)

            # time.sleep(20)
        except Exception as e:
            print(f"Error en la fila {index}: {e}")
            break

    # Cerrar el navegador
    hora_finalizacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Inicio de sesión exitoso, URL actualizada a las {hora_finalizacion}.")
    print("El Bot a Finalizado con exito!, todos los Reguistros PAF")

    driver.quit()