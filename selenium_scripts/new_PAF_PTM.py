# new_PAF_PTM.py
import time
import os
import pandas as pd
from datetime import datetime
from selenium import webdriver
# Importaciones adicionales
from PyQt5.QtWidgets import QFileDialog
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    # Configurar y guardar el horario si no es "-Sin Atención-"
    if tipo_horario != "-Sin Atención-":
        establecer_horario(driver, tipo_horario, horario_dia, prefijo_id)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, f"{prefijo_id}btnAdicionarHorario"))
        ).click()

        # Acciones adicionales para "24 horas" si es necesario
        if horario_dia[1] >= "23:00" or horario_dia[-1] >= "23:00":
            # Esperar y hacer clic en la ventana emergente si es necesario
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, f"{prefijo_id}ASPxPopupControlMensaje_HCB-1"))
            ).click()
# Función para dividir los horarios en horario de mañana y tarde
def dividir_horarios(horario):
    if pd.isnull(horario) or horario in ["0", "00:00", "00:00 a 00:00", "00:00 a 00:00 y 00:00 a 00:00"]:
        return ["0:00", "0:00", "", ""]  # Modificado para manejar "0" y "00:00" como "Sin Atención"
    elif "y" in horario:
        partes = horario.split(' y ')
        manana = partes[0].split(' a ')
        tarde = partes[1].split(' a ')
        return manana + tarde
    else:
        partes = horario.split(' a ')
        return partes + ["", ""]  # Agrega campos vacíos para entrada2 y salida2
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
    webdriver_path = r'C:\msedgedriver\msedgedriver.exe'

    # Ruta al ejecutable de Edge Dev
    edge_dev_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

    # Ruta al perfil de usuario específico de Edge Dev (ajusta esto según sea necesario)
    perfil_usuario = r'C:\Users\tellezd\AppData\Local\Microsoft\Edge\User Data\Default'

    # Configuración de las opciones para apuntar a Edge Dev y al perfil específico
    options = Options()
    options.binary_location = edge_dev_path
    options.add_argument(f'user-data-dir={perfil_usuario}')

    # Crear el servicio con la ruta del WebDriver
    service = Service(webdriver_path)

    # Iniciar Edge Dev con las opciones configuradas
    driver = webdriver.Edge(service=service, options=options)

    # Lee los datos desde el archivo Excel y forza todas las columnas a ser tratadas como texto
    df = pd.read_excel(
        ruta_archivo_excel,
        sheet_name="Cargas_de_Alta_BOT",
        dtype=str  # Forza todas las columnas a ser tratadas como texto
    )

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

            # Horarios
            horario_LUN = dividir_horarios(row["HORARIO_LUNES"])
            horario_MAR = dividir_horarios(row["HORARIO_MARTES"])
            horario_MIE = dividir_horarios(row["HORARIO_MIERCOLES"])
            horario_JUE = dividir_horarios(row["HORARIO_JUEVES"])
            horario_VIE = dividir_horarios(row["HORARIO_VIERNES"])
            horario_SAB = dividir_horarios(row["HORARIO_SABADO"])
            horario_DOM = dividir_horarios(row["HORARIO_DOMINGO"])

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

            # Asignar tipos de horarios y horarios para cada día de la semana
            dias = {
                1: (horario_LUN),
                2: (horario_MAR),
                3: (horario_MIE),
                4: (horario_JUE),
                5: (horario_VIE),
                6: (horario_SAB),
                7: (horario_DOM)
            }

            # Agrupar días con horarios comunes
            grupos_horarios = agrupar_dias_con_horarios_comunes(dias)
            for clave, dias_grupo in grupos_horarios.items():
                tipo_horario, horarios_dia = clave

                # Saltar el proceso de guardar horario si es "-Sin Atención-"
                if tipo_horario == "-Sin Atención-":
                    continue

                # Seleccionar los días del grupo
                seleccionar_dias(driver, dias_grupo)

                # Establecer y guardar el horario
                # if tipo_horario not in ["-Sin Atención-"]:
                #     establecer_horario(driver, tipo_horario, horarios_dia, 'MainContent_DefaultContent_')

                guardar_horario(driver, tipo_horario, horarios_dia, 'MainContent_DefaultContent_')

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
            # Añadir la fila con error a la lista
            registros_con_errores.append(row)

    # Cerrar el navegador
    hora_finalizacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Inicio de sesión exitoso, URL actualizada a las {hora_finalizacion}.")
    print("El Bot a Finalizado con exito!, todos los Reguistros")

    # Guardar los registros con errores
    if registros_con_errores:
        # Puedes descomentar la siguiente línea para permitir que el usuario elija la ruta
        # ruta_archivo_errores = seleccionar_ruta_guardado()
        # O usar la ruta del escritorio por defecto
        ruta_archivo_errores = os.path.join(obtener_ruta_escritorio(), "errores.xlsx")

        if ruta_archivo_errores:  # Verificar si se ha proporcionado una ruta
            df_errores = pd.DataFrame(registros_con_errores)
            with pd.ExcelWriter(ruta_archivo_errores, mode='a', if_sheet_exists='overlay') as writer:
                df_errores.to_excel(writer, sheet_name='Errores', index=False)


    driver.quit()