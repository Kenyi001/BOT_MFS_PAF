import sys
import time
import pandas as pd
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def es_hora_mayor_o_igual(hora, hora_comparacion="23:00"):
    hora_formato = datetime.strptime(hora, "%H:%M")
    hora_comparacion_formato = datetime.strptime(hora_comparacion, "%H:%M")
    return hora_formato >= hora_comparacion_formato
def realizar_acciones_basedo_en_horario(driver, *horarios):
    # Comprobar si alguno de los horarios es mayor o igual a 23:00
    if any(es_hora_mayor_o_igual(hora) for hora in horarios):
        # Código para hacer clic en Guardar y luego en la ventana emergente
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnAdicionarHorario"))
        ).click()
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_ASPxPopupControlMensaje_HCB-1"))
        ).click()
    else:
        # Código para hacer clic en Guardar si los horarios son menores a 23:00
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnAdicionarHorario"))
        ).click()
def seleccionar_dias(driver, dias):
    for dia in dias:
        checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"td:nth-child({dia}) > label")))
        checkbox.click()
def buscar_y_eliminar_horario(driver, dia_buscado):
    # Esperar a que la tabla de horarios esté visible
    tabla_horarios = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_gvHorariosAtencion_DXMainTable"))
    )
    # Encontrar todas las filas en la tabla
    filas = tabla_horarios.find_elements(By.TAG_NAME, "tr")
    # Iterar a través de las filas para encontrar el día buscado
    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")
        if len(celdas) > 1 and dia_buscado in celdas[0].text:
            # Encontrar el botón Borrar en esta fila y hacer clic
            boton_borrar = celdas[-1].find_element(By.TAG_NAME, "a")
            boton_borrar.click()
            return  # Sale de la función después de encontrar y hacer clic en el botón
def establecer_tipo_horario(driver, tipo_horario_id, tipo_horario_text):
    select_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, tipo_horario_id))
    )
    Select(select_element).select_by_visible_text(tipo_horario_text)

def establecer_horario(driver, tipo_horario_text, entrada1, salida1, entrada2, salida2, prefijo_id):
    # Seleccionar tipo de horario
    establecer_tipo_horario(driver, f'{prefijo_id}ddlTipoHorario', tipo_horario_text)

    # Configurar los horarios usando JavaScript
    driver.find_element(By.ID, "MainContent_DefaultContent_timeEntrada1_I").click()
    driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada1_I').value = '{entrada1}';")
    driver.find_element(By.ID, "MainContent_DefaultContent_timeSalida1_I").click()
    driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida1_I').value = '{salida1}';")
    driver.find_element(By.ID, "MainContent_DefaultContent_timeEntrada2_I").click()
    driver.execute_script(f"document.getElementById('{prefijo_id}timeEntrada2_I').value = '{entrada2}';")
    driver.find_element(By.ID, "MainContent_DefaultContent_timeSalida2_I").click()
    driver.execute_script(f"document.getElementById('{prefijo_id}timeSalida2_I').value = '{salida2}';")

def ejecutar_automatizacion_modificacion(ruta_excel):
    try:
        # Configuración del WebDriver de Edge
        webdriver_path = r'C:\msedgedriver\msedgedriver.exe'
        edge_dev_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
        perfil_usuario = r'C:\Users\tellezd\AppData\Local\Microsoft\Edge\User Data\Default'

        options = Options()
        options.binary_location = edge_dev_path
        options.add_argument(f'user-data-dir={perfil_usuario}')
        service = Service(webdriver_path)
        driver = webdriver.Edge(service=service, options=options)

        mensaje_deseado = "Se guardo el registro de Punto de Atención con la asignación"
        # Lee los datos desde el archivo Excel y forza todas las columnas a ser tratadas como texto
        df = pd.read_excel(
            ruta_excel,
            sheet_name="Cargas_de_Alta_BOT",
            dtype=str  # Forza todas las columnas a ser tratadas como texto
        )

        driver.get("https://appweb.asfi.gob.bo/RMI/Login.aspx?ReturnUrl=%2fRMI%2fDefault.aspx")

        try:
            nueva_url = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
            WebDriverWait(driver, 60).until(EC.url_to_be(nueva_url))
            hora_iniciacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"Inicio de sesión exitoso, URL actualizada a las {hora_iniciacion}.")
        except TimeoutException:
            print("Tiempo de espera excedido para el cambio de URL después del inicio de sesión.")

        contador_row = 0
        for index, row in df.iterrows():
            contador_row += 1
            mef = str(row['MEF'])[-4:]
            # nueva_direccion = row['DIRECCIONES']
            # nombre_comercio = row['NOMBRE_PTM']

            # Asignaciones correctas dentro del bucle
            horario_entrada1_SAB = row["HORARIO_SAB_MAÑ_INI"]
            horario_salida1_SAB = row["HORARIO_SAB_MAÑ_FIN"]
            horario_entrada2_SAB = row["HORARIO_SAB_TAR_INI"]
            horario_salida2_SAB = row["HORARIO_SAB_TAR_FIN"]
            horario_entrada1_DOM = row["HORARIO_DOM_MAÑ_INI"]
            horario_salida1_DOM = row["HORARIO_DOM_MAÑ_FIN"]
            horario_entrada2_DOM = row["HORARIO_DOM_TAR_INI"]
            horario_salida2_DOM = row["HORARIO_DOM_TAR_FIN"]
            tipo_de_horario_SAB = row["TIPO_HORARIO_SAB"]
            tipo_de_horario_DOM = row["TIPO_HORARIO_DOM"]

            # Horario Lunes-Viernes, Sábado y Domingo
            horario_sabado = horario_entrada1_SAB + "-" + horario_salida1_SAB + "-" + horario_entrada2_SAB + "-" + horario_salida2_SAB
            horario_domingo = horario_entrada1_DOM + "-" + horario_salida1_DOM + "-" + horario_entrada2_DOM + "-" + horario_salida2_DOM

            driver.get("https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx")

            input_mef = driver.find_element(By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")
            input_mef.clear()
            input_mef.send_keys(mef)
            time.sleep(5)

            # Añadido no Probado
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_btnEditar")))

            boton_editar = driver.find_element(By.ID, "MainContent_DefaultContent_btnEditar")
            boton_editar.click()

            # Verificar y configurar horario para Sábado
            if tipo_de_horario_SAB != "-Sin Atención-":
                buscar_y_eliminar_horario(driver, "Sabado")  # Eliminar horario de sábado
                seleccionar_dias(driver, [6])
                establecer_horario(driver, tipo_de_horario_SAB, horario_entrada1_SAB, horario_salida1_SAB,
                                   horario_entrada2_SAB, horario_salida2_SAB, 'MainContent_DefaultContent_')

                # Llamar a la función para guardar y realizar acciones basadas en el horario
                realizar_acciones_basedo_en_horario(driver, horario_salida2_SAB)

            # Verificar y configurar horario para Domingo
            if tipo_de_horario_DOM != "-Sin Atención-":
                buscar_y_eliminar_horario(driver, "Domingo")  # Eliminar horario de domingo
                if horario_domingo != horario_sabado and tipo_de_horario_SAB != "-Sin Atención-":
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnAdicionarHorario"))).click()
                seleccionar_dias(driver, [7])
                establecer_horario(driver, tipo_de_horario_DOM, horario_entrada1_DOM, horario_salida1_DOM,
                                   horario_entrada2_DOM, horario_salida2_DOM, 'MainContent_DefaultContent_')

                # Llamar a la función para guardar y realizar acciones basadas en el horario
                realizar_acciones_basedo_en_horario(driver, horario_salida2_DOM)

            # campo_direccion = driver.find_element(By.ID, "MainContent_DefaultContent_txtDireccion")
            # campo_direccion.clear()

            # Reemplaza 'ID_del_elemento' con el ID de un elemento que se carga después de la recarga de la página
            # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_txtDireccion")))

            # Identifica el elemento por su ID
            # id_elemento = "MainContent_DefaultContent_txtDireccion"

            # Ejecuta un script de JavaScript para cambiar el valor del elemento
            # driver.execute_script(f"document.getElementById('{id_elemento}').value = '{nueva_direccion}';")

            # Identifica el elemento por su ID
            #id_elemento_nombre_comercio = "MainContent_DefaultContent_txtNombre"

            # Ejecuta un script de JavaScript para cambiar el valor del elemento
            #driver.execute_script(f"document.getElementById('{id_elemento_nombre_comercio}').value = '{nombre_comercio}';")

            # time.sleep(60)
            driver.find_element(By.ID, "MainContent_DefaultContent_btnGrabar").click()

            try:

                # Esperar hasta que el elemento con el mensaje se haga visible
                mensaje_element = WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.ID, "MainContent_DefaultContent_lblMensaje"))
                )
                mensaje_texto = mensaje_element.text[:-6]

                if mensaje_deseado in mensaje_texto:
                    print("El mensaje deseado está presente:", mensaje_texto, " - ", contador_row, " - ", mef)
                else:
                    print("El mensaje deseado no está presente:", mensaje_texto, " - ", contador_row, " - ", mef)
                    print("Deteniendo el bot de modificación.")
                    sys.exit()
            except TimeoutException:
                print(
                    "Tiempo de espera excedido esperando el mensaje después de grabar. Verifica si el ID del elemento es correcto o si la página está tomando demasiado tiempo para responder.")
                sys.exit()
    except TimeoutException as e:
        print(f"Error: Tiempo de espera excedido - {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

    finally:
        hora_finalizacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"El bot finalizó (con éxito o error) a las {hora_finalizacion}.")
        driver.quit()

    driver.quit()