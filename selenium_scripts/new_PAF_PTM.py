# new_PAF_PTM.py
import os
from datetime import datetime
from tkinter import messagebox

import pandas as pd
import openpyxl
import tk
# Importaciones adicionales
from PyQt5.QtWidgets import QFileDialog, QApplication, QMessageBox
from selenium import webdriver
from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

ruta_actual = os.getcwd()
def mostrar_ventana_emergente():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Critical)  # Establecer el icono a rojo
    msgBox.setWindowTitle("Notificacion de Error")
    msgBox.setText("Se generaron algunos errores en la Carga de los MEF y se guardaron en un archivo Excel.")
    msgBox.exec()

def esperar_y_encontrar_elemento(driver, by, identificador, tiempo_espera=20):
    """Espera hasta que un elemento sea localizable y luego lo retorna."""
    return WebDriverWait(driver, tiempo_espera).until(
        EC.presence_of_element_located((by, identificador))
    )

def iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena):
    """Inicia sesión en un sitio web proporcionado utilizando credenciales de usuario."""
    driver.get(url_inicio_sesion)

    # Esperar hasta que los campos de usuario y contraseña estén presentes
    campo_usuario = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_txtUsuario")
    campo_contrasena = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_txtPassword")

    # Llenar y enviar el formulario de inicio de sesión
    boton_inicio_sesion = esperar_y_encontrar_elemento(driver, By.ID, "MainContent_DefaultContent_LoginButton")

    campo_usuario.send_keys(usuario)
    campo_contrasena.send_keys(contrasena)
    boton_inicio_sesion.click()

def establecer_coordenadas(driver, coordenada_x, coordenada_y, id_coordenada_x, id_coordenada_y):
    """
    Habilita y establece las coordenadas X e Y en los campos especificados mediante ID.

    Args:
        driver: Instancia del navegador usado por Selenium.
        coordenada_x (str): Valor de la coordenada X a establecer.
        coordenada_y (str): Valor de la coordenada Y a establecer.
        id_coordenada_x (str): ID del campo de entrada de la coordenada X.
        id_coordenada_y (str): ID del campo de entrada de la coordenada Y.
    """
    # Espera que los elementos estén presentes antes de interactuar con ellos.
    esperar_y_encontrar_elemento(driver, By.ID, id_coordenada_x)
    esperar_y_encontrar_elemento(driver, By.ID, id_coordenada_y)

    # Habilitar los campos de coordenadas y establecer sus valores.
    script = (
        f"document.getElementById('{id_coordenada_x}').removeAttribute('disabled');"
        f"document.getElementById('{id_coordenada_y}').removeAttribute('disabled');"
        f"document.getElementById('{id_coordenada_x}').value = '{coordenada_x}';"
        f"document.getElementById('{id_coordenada_y}').value = '{coordenada_y}';"
    )
    driver.execute_script(script)
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
    """
    Selecciona una localidad y departamento en una tabla dada, basado en los valores buscados.

    Args:
        driver: El driver de Selenium.
        tabla_id (str): El ID de la tabla HTML donde buscar.
        localidad_buscada (str): La localidad que se busca seleccionar.
        departamento_verificado (str): El departamento asociado a la localidad.
    """
    tabla = esperar_y_encontrar_elemento(driver, tabla_id, By.ID)

    # Normalizar texto para la comparación
    localidad_buscada = localidad_buscada.strip().lower()
    departamento_verificado = departamento_verificado.strip().lower()

    filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]  # Excluir la cabecera de la tabla
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
def guardar_formulario_y_verificar_mensaje(driver, boton_guardar_id, mensaje_id, mensaje_deseado, contador_row, numero_MEF, tiempo_espera=10):
    """
    Guarda el formulario y verifica si el mensaje deseado está presente.

    Args:
        driver: El driver de Selenium.
        boton_guardar_id (str): El ID del botón de guardar.
        mensaje_id (str): El ID del elemento que contiene el mensaje de confirmación.
        mensaje_deseado (str): El mensaje esperado para confirmar el éxito.
        contador_row (int): El contador de fila actual para el registro.
        numero_MEF (str): El número MEF del registro actual.
        tiempo_espera (int): El tiempo máximo de espera para la aparición del mensaje.

    Returns:
        bool: True si el mensaje deseado está presente, False en caso contrario.
    """
    try:
        # Hacer clic en el botón de guardar
        driver.find_element(By.ID, boton_guardar_id).click()

        # Esperar hasta que el elemento con el mensaje se haga visible
        mensaje_element = WebDriverWait(driver, tiempo_espera).until(
            EC.visibility_of_element_located((By.ID, mensaje_id))
        )

        # Obtener el texto del mensaje
        mensaje_texto = mensaje_element.text

        # Verificar si el mensaje deseado está presente en el texto
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
    # Configuración inicial para modo headless
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')  # Si necesitas cargar un perfil de usuario específico
    service = Service(WEBDRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)

    print("Iniciando el navegador Edge en modo headless...")

    # URL de inicio de sesión
    iniciar_sesion(driver, url_inicio_sesion, usuario, contrasena)

    # Asegúrate de ajustar el ID del elemento según tu aplicación para confirmar un inicio de sesión exitoso
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "LoginName1")))
    print("Inicio de sesión exitoso.")

    # Aquí continúa el procesamiento de tu archivo Excel y la automatización subsecuente
    try:
        df = pd.read_excel(
            ruta_archivo_excel,
            sheet_name=nombre_hoja,  # Nombre de la Hoja de Cálculo
            dtype=str  # Forzar todas las columnas a ser tratadas como texto
        )
        datos_extraidos = []
        print("Lectura de archivo Excel exitosa.")
    except Exception as e:
        print(f"Error al abrir el archivo Excel: {e}")
        return  # Finaliza la ejecución si hay un error al leer el archivo
    contador_row = 0
    ultima_accion = ""
    for index, row in df.iterrows():
        contador_row += 1
        # Accede a los valores de las columnas que necesitas
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
            ultima_accion = "Estableciendo Tipo de Sucursal"
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlTipoSucursal'))
            select.select_by_visible_text(tipo_sucursal)

            # Esperar para visualizar la página
            driver.implicitly_wait(30)  # O ajusta este tiempo según sea necesario

            # Interacción con campos de texto y selección de opciones
            ultima_accion = "Estableciendo Nombre y Cédula de Responsable"
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNombreResponsablePCNF").send_keys(nombre_responsable)
            driver.find_element(By.ID, "MainContent_DefaultContent_txtIdentificacionPCNF").send_keys(identificacion_responsable)

            # espera que se cargue el elemento Nombre de Comercio
            ultima_accion = "Estableciendo Nombre de Comercio"
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_txtNombre")))
            driver.find_element(By.ID, "MainContent_DefaultContent_txtNombre").send_keys(nombre_comercio)

            # espera que se cargue el elemento descripcion de direccion
            # Identifica el elemento por su ID
            id_elemento = "MainContent_DefaultContent_txtDireccion"


            # Ejecuta un script de JavaScript para cambiar el valor del elemento
            ultima_accion = "Estableciendo Direccion"
            driver.execute_script(f"document.getElementById('{id_elemento}').value = '{direccion}';")

            # Interactuar con dropdowns
            select = Select(driver.find_element(By.ID, 'MainContent_DefaultContent_ddlNivelSeguridad'))
            select.select_by_visible_text(nivel_seguridad)

            # Interacción con campos de teléfono y fax
            ultima_accion = "Estableciendo Telefono y fax"
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtTelefono_I").send_keys(telefono)
            driver.find_element(By.ID, "ctl00_ctl00_MainContent_DefaultContent_ASPxtxtFax_I").send_keys(fax)

            # Interaccion con el campo de Número de Cámaras de Seguridad
            # Localizar el campo de entrada de cámaras de seguridad por ID
            ultima_accion = "Estableciendo Camaras de Seguridad"
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
            ultima_accion = "Estableciendo Horarios"
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
            ultima_accion = "Estableciendo Servicios"
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn0_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn1_D").click()
            driver.find_element(By.ID, "MainContent_DefaultContent_ASPxGridViewServicios_DXSelBtn2_D").click()

            # Uso de las funciones
            ultima_accion = "Estableciendo Localidad"
            localidad_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_txtLocalidad"
            boton_buscar_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_btnBuscar"
            tabla_id = "MainContent_DefaultContent_ctlPaisLocalidadGeografia2012_gridGeografiaLocalidades"

            buscar_localidad(driver, localidad_id, boton_buscar_id, localidad_busqueda)
            seleccionar_localidad(driver, tabla_id, localidad_busqueda, departamento_verificado)

            # Habilitar y llenar campos de coordenadas
            ultima_accion = "Estableciendo coordenadas"
            id_coordenada_x = "MainContent_DefaultContent_txtCoordenadaX"
            id_coordenada_y = "MainContent_DefaultContent_txtCoordenadaY"
            establecer_coordenadas(driver, coordenada_x, coordenada_y, id_coordenada_x, id_coordenada_y)

            # Guardar el formulario y verificar el mensaje
            resultado_guardado = guardar_formulario_y_verificar_mensaje(
                driver,
                "MainContent_DefaultContent_btnGrabar",
                "MainContent_DefaultContent_lblMensaje",
                mensaje_deseado,
                contador_row,
                numero_MEF
            )

            # Tomar alguna acción basada en el resultado
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
    if datos_extraidos:
        ruta_carpeta = os.path.join(ruta_actual, 'Errores de Carga')
        if not os.path.exists(ruta_carpeta):
            os.makedirs(ruta_carpeta)
        formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
        nombre_archivo = f"PTM_Error_Carga_{formato_fecha_hora}.xlsx"
        nombre_archivo_salida = os.path.join(ruta_carpeta, nombre_archivo)
        df_extraido = pd.DataFrame(datos_extraidos)
        df_extraido.to_excel(nombre_archivo_salida, index=False)

        # Crear una ventana emergente
        mostrar_ventana_emergente()


    # Cerrar el navegador
    driver.quit()