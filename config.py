# config.py
import os

# Configuración de las rutas para el WebDriver y el perfil de usuario
WEBDRIVER_PATH = os.getenv('EDGE_PATH')
NAV_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH')
USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH')
URL_INICIO_SESION = 'https://appweb.asfi.gob.bo/RMI/Default.aspx'

# Asegúrate de que las variables no estén vacías
if not all([WEBDRIVER_PATH, NAV_BINARY_PATH, USER_PROFILE_PATH]):
    # Configuración de las rutas para el WebDriver y el perfil de usuario
    WEBDRIVER_PATH = r'C:\msedgedriver\msedgedriver.exe'
    NAV_BINARY_PATH = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
    USER_PROFILE_PATH = r'C:\Users\Tellezd\AppData\Local\Microsoft\Edge\User Data\Default'

    # raise ValueError("Una o más variables de entorno requeridas no están definidas.")
