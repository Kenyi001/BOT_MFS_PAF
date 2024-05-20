# config.py
import os

# Configuración de las rutas para el WebDriver y el perfil de usuario
WEBDRIVER_PATH = os.getenv('EDGE_PATH')
EDGE_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH')
USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH')
URL_INICIO_SESION = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
URL_REGISTRO_PAF = "https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx"
