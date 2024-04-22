# config.py
import os

WEBDRIVER_PATH = r'C:\msedgedriver\msedgedriver.exe'
EDGE_BINARY_PATH = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
USER_PROFILE_PATH = r'C:\Users\Kenji\AppData\Local\Microsoft\Edge\User Data\Default'
URL_INICIO_SESION = 'https://appweb.asfi.gob.bo/RMI/Default.aspx'

# # Configuración de las rutas para el WebDriver y el perfil de usuario
# WEBDRIVER_PATH = os.getenv('EDGE_PATH')
# EDGE_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH')
# USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH')
# URL_INICIO_SESION = 'https://appweb.asfi.gob.bo/RMI/Default.aspx'
#
# # Asegúrate de que las variables no estén vacías
# if not all([WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH]):
#     # Configuración de las rutas para el WebDriver y el perfil de usuario
#     WEBDRIVER_PATH = r'"C:\msedgedriver\msedgedriver.exe"'
#     EDGE_BINARY_PATH = r'"C:\Program Files (x86)\Microsoft\EdgeCore\123.0.2420.97\msedge.exe"'
#     USER_PROFILE_PATH = r'C:\Users\Kenji\AppData\Local\Microsoft\Edge\User Data\Default'

    # raise ValueError("Una o más variables de entorno requeridas no están definidas.")
