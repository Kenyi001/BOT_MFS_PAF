# config.py
import os

# Configuración de las rutas para el WebDriver y el perfil de usuario
WEBDRIVER_PATH = os.getenv('EDGE_PATH')
EDGE_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH')
USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH')

print(f"WebDriver Path: {WEBDRIVER_PATH}")
print(f"Edge Binary Path: {EDGE_BINARY_PATH}")
print(f"User Profile Path: {USER_PROFILE_PATH}")

# Asegúrate de que las variables no estén vacías
if not all([WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH]):
    raise ValueError("Una o más variables de entorno requeridas no están definidas.")
