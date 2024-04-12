import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog

# Importa tus scripts de automatización
from selenium_scripts.new_PAF_PTM import ejecutar_automatizacion_new
from selenium_scripts.modificacion_PAF_PTM import ejecutar_automatizacion_modificacion
from selenium_scripts.extraccion_datos_web import extraer_datos_web  # Importa la función de extracción de datos web

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Selector de Automatización y Archivo')
        self.resize(300, 100)

        layout = QVBoxLayout()

        # Botón para la automatización de New PAF/PTM
        self.btnNewPAFPTM = QPushButton('New PAF/PTM')
        self.btnNewPAFPTM.clicked.connect(lambda: self.openFileDialog(self.ejecutar_new_PAF_PTM))
        layout.addWidget(self.btnNewPAFPTM)

        # Botón para la automatización de Modificación PAF/PTM
        self.btnModificacionPAFPTM = QPushButton('Modificación PAF/PTM')
        self.btnModificacionPAFPTM.clicked.connect(lambda: self.openFileDialog(self.ejecutar_modificacion_PAF_PTM))
        layout.addWidget(self.btnModificacionPAFPTM)

        # Botón para la extracción de datos web
        self.btnExtraccionDatosWeb = QPushButton('Extracción de Datos Web')
        self.btnExtraccionDatosWeb.clicked.connect(lambda: self.openFileDialog(self.ejecutar_extraccion_datos_web))
        layout.addWidget(self.btnExtraccionDatosWeb)

        self.setLayout(layout)

    def openFileDialog(self, automation_function):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel",
                                              "",  # Aquí se debe dejar vacío para abrir en el último directorio visitado o en el predeterminado
                                              "Excel Files (*.xlsx)", options=options)
        # Si el usuario cancela la selección de archivo, no hagas nada
        if not file:
            return

        automation_function(file)

    def ejecutar_new_PAF_PTM(self, file):
        ejecutar_automatizacion_new(file)

    def ejecutar_modificacion_PAF_PTM(self, file):
        ejecutar_automatizacion_modificacion(file)

    def ejecutar_extraccion_datos_web(self, file):
        extraer_datos_web("https://appweb.asfi.gob.bo/RMI/Default.aspx", file)  # Reemplaza con la URL correcta

app = QApplication(sys.argv)
demo = AppDemo()
demo.show()
sys.exit(app.exec_())
