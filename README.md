# BOT de Automatizacion ASFI

Un bot con la capacidad de realizar la *Subida* y *Modificacion* de datos de los PAF (PTM) en cara a la plataforma de la ASFI
, envia formulario (Alta PTM) y modifica datos de los PAF (cambio de datos de un PTM).

## Comenzando

Estas instrucciones te proporcionarán una copia del proyecto en funcionamiento en tu máquina local para propósitos de desarrollo y pruebas.

### Prerrequisitos

Lo que necesitas para instalar el software y cómo instalarlos.

python >= 3.6 <br> pip >= 19.0

### Instalación

Una serie de ejemplos paso a paso que te dicen cómo hacer un entorno de desarrollo en ejecución.

Clona el repositorio:

```bash
git clone [URL del repositorio]

cd [nombre del repositorio]
```
Instala las dependencias:

```bash
pip install -r requirements.txt
```

### Configuración

Para que el bot funcione correctamente, necesitas configurar las siguientes variables de entorno:

Reemplaza `[Ruta del webdriver]`, `[Ruta del ejecutable de Edge]` y `[Ruta del perfil de usuario]` con las rutas correspondientes en tu sistema.

en caso de que no tengas el perfil de usuario, puedes crearlo con el siguiente comando:

```bash
msedge.exe --profile-directory="Profile 1"
```

```bash
# Ruta del webdriver
export WEBDRIVER_PATH=[Ruta del webdriver]

# Ruta del ejecutable de Edge
export EDGE_PATH=[Ruta del ejecutable de Edge]

# Ruta del perfil de usuario
export USER_DATA_DIR=[Ruta del perfil de usuario]
```

### Ejecución

Para ejecutar el bot, ejecuta el siguiente comando:

```bash
python main.py
```

### Luego configura las variables en el archivo .env

Ejecución de pruebas
Explica cómo ejecutar las pruebas automatizadas para este sistema.

### Ejemplo
python -m unittest discover

### Despliegue
Agrega notas adicionales sobre cómo desplegar esto en un sistema en vivo.

### Construido con
Python - El lenguaje de programación usado

### Contribuir
Por favor, lee CONTRIBUTING.md para detalles sobre nuestro código de conducta, y el proceso para enviarnos pull requests.

### Versionado
Usamos SemVer para el versionado. Para las versiones disponibles, mira las tags en este repositorio.

### Autores
Dax Kenji Tellez Duran - Tellezd
Licencia
Este proyecto está licenciado bajo la Licencia XYZ - mira el archivo LICENSE.md para detalles.

### Reconocimientos
Hat tip a cualquier persona cuyo código fue usado
Inspiración
etc

Recuerda cambiar los marcadores de posición como "[URL del repositorio]", "[TuUsuario]", y cualquier otra información específica por los detalles reales de tu proyecto. Esto es solo un punto de partida; ajusta y expande cada sección para adaptarse a las necesidades y la complejidad de tu proyecto.

