import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog

# Importa tus scripts de automatización
from selenium_scripts.new_PAF_PTM import ejecutar_automatizacion_new
from selenium_scripts.modificacion_PAF_PTM import ejecutar_automatizacion_modificacion

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Selector de Automatización y Archivo')
        self.resize(300, 100)

        layout = QVBoxLayout()

        # Botón para la automatización de New PAF/PTM
        self.btnNewPAFPTM = QPushButton('New PAF/PTM')
        self.btnNewPAFPTM.clicked.connect(lambda: self.openFileDialog(ejecutar_automatizacion_new))
        layout.addWidget(self.btnNewPAFPTM)

        # Botón para la automatización de Modificación PAF/PTM
        self.btnModificacionPAFPTM = QPushButton('Modificación PAF/PTM')
        self.btnModificacionPAFPTM.clicked.connect(lambda: self.openFileDialog(ejecutar_automatizacion_modificacion))
        layout.addWidget(self.btnModificacionPAFPTM)

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

app = QApplication(sys.argv)
demo = AppDemo()
demo.show()
sys.exit(app.exec_())
