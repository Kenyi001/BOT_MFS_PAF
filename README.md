# BOT PTM - ASFI

Esta aplicación está diseñada para automatizar tareas específicas relacionadas con la gestión de Puntos de Transferencia Monetaria (PTM) y la extracción de datos web. Ofrece una interfaz gráfica amigable que permite al usuario interactuar fácilmente con sus funcionalidades.

## Comenzando

Estas instrucciones te proporcionarán una copia del proyecto en funcionamiento en tu máquina local para propósitos de desarrollo y pruebas.

### Prerrequisitos

Lo que necesitas para instalar el software:

- Python 3.7+
- pip (gestor de paquetes de Python)
- Navegador Microsoft Edge
- Microsoft Edge WebDriver

se desarrollo en el ide de **PyCharm 2024.1**
### Instalación

1. Clona el repositorio en tu máquina local:

   ```sh
   git clone https://github.com/tu_usuario/tu_proyecto.git
cd tu_proyecto
2. Crea y activa un entorno virtual:

   ```sh
   python -m venv venv
source venv/bin/activate  # En Windows usa `venv\Scripts\activate`
3. Instala las dependencias necesarias:

   ```sh
   pip install -r requirements.txt
4. Configura los archivos de configuración:

Asegúrate de tener config.py con las rutas correctas para 

WEBDRIVER_PATH, EDGE_BINARY_PATH, y USER_PROFILE_PATH.

Coloca el WebDriver de Edge en la ruta especificada en config.py.

### Uso
1. Para ejecutar la aplicación, usa el siguiente comando:
   ```sh
   python main2.py

La aplicación te pedirá las rutas y credenciales necesarias para la automatización.

### Funcionalidades
1. Solicitar Rutas: Permite al usuario configurar las rutas necesarias para el funcionamiento del bot.

2. Automatización de Nuevo PAF/PTM: Realiza la creación de nuevos registros de PTM basándose en un archivo Excel proporcionado por el usuario. 

   Los campos predeterminados son:
   - Tipo de Sucursal: Punto de Atención Corresponsal No Financiero
   - Nivel de Seguridad: Bajo
   - Fax: 0
   - Coordenadas: lat_gps y lng_gps
   - Horario de Atención: 
     - 00:00 a 00:00 son sin atencion
     - 00:01 a 00:01 son 24 hrs
     

3. Extracción de Datos Web: Extrae datos específicos desde una web utilizando la información proporcionada en un archivo Excel.
   Ejecución de pruebas

4. Para ejecutar las pruebas automatizadas, usa:

   ```sh
   pytest
Asegúrate de que tienes las dependencias necesarias instaladas y que el WebDriver está configurado correctamente.

### Despliegue
Para crear un ejecutable de la aplicación:

1. Instala PyInstaller:

   ```sh
   pip install pyinstaller
   
2. Ejecuta PyInstaller con el siguiente comando:

   ```sh
   pyinstaller --onefile --windowed --icon=\you\path\BOT_MFS_PAF\Static\Tigo Money logo.webp --name=Bot_PTM_Asfi main2.py
Esto generará un archivo ejecutable que puedes distribuir.

### Construido con
- Python - El lenguaje de programación usado.
- PyQt5 - Librería para la creación de la interfaz gráfica.
- Selenium - Para la automatización de navegadores web.
- Pandas - Para la manipulación de datos.
- openpyxl - Para trabajar con archivos Excel.

### Autores
Dax Kenji Tellez Duran - todo el proyecto - Tellezd

### Reconocimientos
Quisiera expresar mi más profundo 
agradecimiento a Daniel, cuyo apoyo incondicional 
ha sido fundamental en este proyecto. 
Su confianza en mis capacidades me ha permitido dar 
el primer paso y avanzar con seguridad. También quiero agradecer a Marcos Salome, cuyas propuestas de mejora siempre han sido valiosas y han enriquecido mi trabajo. 
Finalmente, mi gratitud a mi tutor Mauricio Perico por brindarme la oportunidad de estar en su equipo y por su orientación constante. Su apoyo ha sido invaluable en este viaje. Gracias a todos.