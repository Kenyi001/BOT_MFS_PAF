# config.py
import os

# Configuración de las rutas para el WebDriver y el perfil de usuario
WEBDRIVER_PATH = os.getenv('EDGE_PATH')
EDGE_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH')
USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH')
URL_INICIO_SESION = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
URL_REGISTRO_PAF = "https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx"

# Asegúrate de que las variables no estén vacías
if not all([WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH]):
    # Configuración de las rutas para el WebDriver y el perfil de usuario
    WEBDRIVER_PATH = r''
    EDGE_BINARY_PATH = r''
    USER_PROFILE_PATH = r''

    # raise ValueError("Una o más variables de entorno requeridas no están definidas.")
