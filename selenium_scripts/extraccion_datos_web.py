import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH

# Variables para el usuario y contraseña
def buscar_y_editar_mef(driver, mef):
    # Ingresar el MEF en el campo de búsqueda
    input_mef = driver.find_element(By.ID, "MainContent_DefaultContent_gridPuntoAtencion1_DXFREditorcol3_I")
    input_mef.clear()
    input_mef.send_keys(mef)
    time.sleep(3)


    # Esperar a que el botón de edición esté visible
    wait = WebDriverWait(driver, 20)
    # Esperar a que el botón de edición esté interactuable
    wait.until(EC.element_to_be_clickable((By.ID, "MainContent_DefaultContent_btnEditar")))

    # Hacer clic en el botón de edición
    boton_editar = driver.find_element(By.ID, "MainContent_DefaultContent_btnEditar")
    boton_editar.click()

def extraer_datos_web(url, archivo_excel, usuario, contrasena):
    # Configuración del WebDriver de Edge para modo headless
    options = Options()
    options.binary_location = EDGE_BINARY_PATH
    options.add_argument(f'user-data-dir={USER_PROFILE_PATH}')
    options.add_argument("--headless")  # Ejecutar en modo headless
    options.add_argument("--disable-gpu")  # Desactivar el GPU
    service = Service(WEBDRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)

    try:
        # Navegar a la página de inicio de sesión
        driver.get(url)

        # Esperar a que los elementos de inicio de sesión estén presentes
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtUsuario")))
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtPassword")))

        # Ingresar usuario y contraseña
        campo_usuario = driver.find_element(By.ID, "MainContent_DefaultContent_txtUsuario")
        campo_contrasena = driver.find_element(By.ID, "MainContent_DefaultContent_txtPassword")

        campo_usuario.send_keys(usuario)
        campo_contrasena.send_keys(contrasena)

        # Hacer clic en el botón de iniciar sesión
        boton_inicio_sesion = driver.find_element(By.ID, "MainContent_DefaultContent_LoginButton")
        boton_inicio_sesion.click()

        # Esperar a que la sesión se inicie correctamente, ajusta esto según tu necesidad
        # Por ejemplo, esperar a una URL específica o a un elemento que solo esté presente tras iniciar sesión
        # WebDriverWait(driver, 240).until(EC.url_to_be("URL_DESEADA_POST_INICIO_SESION"))

        # Leer el archivo Excel
        df = pd.read_excel(archivo_excel, sheet_name="Encontrar_direccion", dtype=str)

        datos_extraidos = []

        # Iterar sobre cada fila del DataFrame
        for index, row in df.iterrows():
            mef_completo = str(row['MEF'])  # Obtener el MEF completo de la fila
            primeros_tres = mef_completo[:3]  # Obtener los primeros tres dígitos
            ultimos_cuatro = mef_completo[-4:]  # Obtener los últimos cuatro dígitos
            mef = f"{primeros_tres}-{ultimos_cuatro}"  # Combinar los dígitos con un guion "-"
            direccion = ""

            driver.get("https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx")
            try:
                # Buscar y editar el MEF
                buscar_y_editar_mef(driver, mef)

                # Esperar hasta que la dirección se cargue
                direccion_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "MainContent_DefaultContent_txtDireccion"))
                )
                direccion = direccion_element.get_attribute("value")
            except TimeoutException:
                print(f"No se pudo extraer la dirección para el MEF {mef}.")
            except Exception as e:
                print(f"Ocurrió un error inesperado: {str(e)}")

            # Agregar los datos extraídos a la lista
            datos_extraidos.append({'MEF': mef, 'Dirección': direccion})

        # Convertir la lista de datos extraídos en un DataFrame
        df_extraido = pd.DataFrame(datos_extraidos)

        # Generar el nombre del archivo con formato de fecha y hora
        formato_fecha_hora = datetime.now().strftime("%Y%m%d_%H%M")
        # Define la ruta de almacenamiento en la parte superior de tu script
        RUTA_ALMACENAMIENTO = "data/Reports"

        # Luego, puedes usar esta variable al generar el nombre del archivo de salida
        nombre_archivo_salida = f"{RUTA_ALMACENAMIENTO}/datos_extraidos_{formato_fecha_hora}.xlsx"

        # Guardar los datos extraídos en un nuevo archivo Excel
        df_extraido.to_excel(nombre_archivo_salida, index=False)

        #print(f"Datos extraídos guardados en '{nombre_archivo_salida}'.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {str(e)}")

    finally:
        return nombre_archivo_salida
        # Cerrar el navegador al finalizar
        driver.quit()

