import os
import sys
import json
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLineEdit, QLabel, \
    QHBoxLayout, QMessageBox, QInputDialog

# Importa tus scripts de automatización
from selenium_scripts.new_PAF_PTM import ejecutar_automatizacion_new
from selenium_scripts.modificacion_PAF_PTM import ejecutar_automatizacion_modificacion
from selenium_scripts.extraccion_datos_web import extraer_datos_web  # Importa la función de extracción de datos web
from config import URL_INICIO_SESION
class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Selector de Automatización y Archivo')
        self.resize(300, 100)

        layout = QVBoxLayout()

        # Botón para solicitar las rutas
        self.btnRequestPaths = QPushButton('Solicitar rutas')
        self.btnRequestPaths.clicked.connect(self.openRequestPathsWindow)
        layout.addWidget(self.btnRequestPaths)

        # Botón para la automatización de New PAF/PTM
        self.btnNewPAFPTM = QPushButton('New PAF/PTM')
        self.btnNewPAFPTM.clicked.connect(lambda: self.openFileDialog(self.ejecutar_new_PAF_PTM, True))
        layout.addWidget(self.btnNewPAFPTM)

        # Botón para la extracción de datos web
        self.btnExtraccionDatosWeb = QPushButton('Extracción de Datos Web')
        self.btnExtraccionDatosWeb.clicked.connect(lambda: self.openFileDialog(self.ejecutar_extraccion_datos_web, True))
        layout.addWidget(self.btnExtraccionDatosWeb)

        self.setLayout(layout)
    def requestCredentials(self):
        username, ok = QInputDialog.getText(self, 'Solicitar credenciales', 'Introduce tu nombre de usuario:', QLineEdit.Normal)
        password, ok = QInputDialog.getText(self, 'Solicitar credenciales', 'Introduce tu contraseña:', QLineEdit.Password)
        if ok and username and password:
            return username, password
        else:
            return None, None

    def openRequestPathsWindow(self):
        self.request_window = QWidget()
        self.request_window.setWindowTitle('Configurar Rutas')
        self.request_window.resize(600, 300)
        layout = QVBoxLayout(self.request_window)

        # Verificar si existe el archivo de configuración y cargarlo
        if os.path.exists('config.json'):
            with open('config.json', 'r') as config_file:
                config = json.load(config_file)
            webdriver_path = config.get('WEBDRIVER_PATH', '')
            edgebinary_path = config.get('EDGE_BINARY_PATH', '')
            userprofile_path = config.get('USER_PROFILE_PATH', '')
        else:
            webdriver_path, edgebinary_path, userprofile_path = '', '', ''

        # Layout para WebDriver Path
        webdriver_layout = QHBoxLayout()
        self.webdriver_input = QLineEdit(webdriver_path)
        webdriver_button = QPushButton('Buscar')
        webdriver_button.clicked.connect(lambda: self.openFileDialog(self.webdriver_input))
        webdriver_layout.addWidget(QLabel('WebDriver Path:'))
        webdriver_layout.addWidget(self.webdriver_input)
        webdriver_layout.addWidget(webdriver_button)
        layout.addLayout(webdriver_layout)

        # Layout para Edge Binary Path
        edgebinary_layout = QHBoxLayout()
        self.edgebinary_input = QLineEdit(edgebinary_path)
        edgebinary_button = QPushButton('Buscar')
        edgebinary_button.clicked.connect(lambda: self.openFileDialog(self.edgebinary_input, True))
        edgebinary_layout.addWidget(QLabel('Edge Binary Path:'))
        edgebinary_layout.addWidget(self.edgebinary_input)
        edgebinary_layout.addWidget(edgebinary_button)
        layout.addLayout(edgebinary_layout)

        # Layout para User Profile Path
        userprofile_layout = QHBoxLayout()
        self.userprofile_input = QLineEdit(userprofile_path)
        userprofile_button = QPushButton('Buscar')
        userprofile_button.clicked.connect(lambda: self.openDirectoryDialog(self.userprofile_input))
        userprofile_layout.addWidget(QLabel('User Profile Path:'))
        userprofile_layout.addWidget(self.userprofile_input)
        userprofile_layout.addWidget(userprofile_button)
        layout.addLayout(userprofile_layout)

        # Botón Aceptar para guardar la configuración
        accept_button = QPushButton('Aceptar')
        accept_button.clicked.connect(self.saveConfiguration)
        layout.addWidget(accept_button)

        self.request_window.setLayout(layout)
        self.request_window.show()

    def openFileDialog(self, line_edit, is_file=False):
        if is_file:
            path, _ = QFileDialog.getOpenFileName(self.request_window, 'Seleccionar Archivo', '', 'All Files (*)')
        else:
            path = QFileDialog.getExistingDirectory(self.request_window, 'Seleccionar Directorio')
        if path:
            line_edit.setText(path)

    def openDirectoryDialog(self, line_edit):
        path = QFileDialog.getExistingDirectory(self.request_window, 'Seleccionar Directorio')
        if path:
            line_edit.setText(path)

    def saveConfiguration(self):
        config = {
            'WEBDRIVER_PATH': self.webdriver_input.text(),
            'EDGE_BINARY_PATH': self.edgebinary_input.text(),
            'USER_PROFILE_PATH': self.userprofile_input.text()
        }
        with open('config.json', 'w') as config_file:
            json.dump(config, config_file)
        QMessageBox.information(self.request_window, "Configuración Guardada", "La configuración ha sido guardada exitosamente.")
        self.request_window.close()
    def ejecutar_new_PAF_PTM(self, file):
        username, password = self.requestCredentials()
        if username and password:
            ejecutar_automatizacion_new(URL_INICIO_SESION, file, username, password)
            self.mostrar_mensaje_confirmacion("La automatización 'New PAF/PTM' se completó con éxito.")

    def ejecutar_modificacion_PAF_PTM(self, file):
        username, password = self.requestCredentials()
        if username and password:
            ejecutar_automatizacion_modificacion(file)
            self.mostrar_mensaje_confirmacion("La automatización 'Modificación PAF/PTM' se completó con éxito.")

    def ejecutar_extraccion_datos_web(self, file):
        username, password = self.requestCredentials()
        if username and password:
            extraer_datos_web(URL_INICIO_SESION, file, username, password)  # Reemplaza con la URL correcta
            self.mostrar_mensaje_confirmacion("La extracción de datos web se completó con éxito.")

    def mostrar_mensaje_confirmacion(self, mensaje):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText(mensaje)
        msgBox.setWindowTitle("Confirmación")
        msgBox.setStandardButtons(QMessageBox.Ok)
        returnValue = msgBox.exec()
        if returnValue == QMessageBox.Ok:
            print('OK clicked')

app = QApplication(sys.argv)
demo = AppDemo()
demo.show()
sys.exit(app.exec_())
