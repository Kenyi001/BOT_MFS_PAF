# config.py
import os
import json

def load_config():
    """Carga la configuración desde un archivo JSON en la carpeta 'resources'."""
    default_config = {"WEBDRIVER_PATH": "", "EDGE_BINARY_PATH": "", "USER_PROFILE_PATH": ""}
    resources_dir = 'resources'
    config_path = os.path.join(resources_dir, 'config.json')

    # Crear la carpeta 'resources' si no existe
    if not os.path.exists(resources_dir):
        os.makedirs(resources_dir)

    # Crear el archivo 'config.json' con la configuración predeterminada si no existe
    if not os.path.exists(config_path):
        with open(config_path, 'w') as config_file:
            json.dump(default_config, config_file, indent=4)

    # Cargar la configuración desde el archivo 'config.json'
    with open(config_path, 'r') as config_file:
        return json.load(config_file)

# Cargar configuración desde el archivo JSON
config = load_config()

# Configuración de las rutas para el WebDriver y el perfil de usuario
WEBDRIVER_PATH = os.getenv('EDGE_PATH', config.get('WEBDRIVER_PATH'))
EDGE_BINARY_PATH = os.getenv('EDGE_DRIVER_PATH', config.get('EDGE_BINARY_PATH'))
USER_PROFILE_PATH = os.getenv('EDGE_PROFILE_PATH', config.get('USER_PROFILE_PATH'))
URL_INICIO_SESION = "https://appweb.asfi.gob.bo/RMI/Default.aspx"
URL_REGISTRO_PAF = "https://appweb.asfi.gob.bo/RMI/RegistroParticipante/puntoAtencion.aspx"
