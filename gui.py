import os
import json
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox, QSizePolicy, QSpacerItem
)

class ConfigurationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Configurar Rutas')
        self.resize(600, 200)

        # Cargar la configuración existente, si la hay
        self.config = self.load_configuration()

        layout = QVBoxLayout(self)

        # Configurar rutas
        self.add_path_layout(layout, 'WebDriver Path:', 'WEBDRIVER_PATH', self.show_info_wp)
        self.add_path_layout(layout, 'Edge Binary Path:', 'EDGE_BINARY_PATH', self.show_info_ep)
        self.add_path_layout(layout, 'User Profile Path:', 'USER_PROFILE_PATH', self.show_info_up)

        # Espaciador para empujar los botones hacia abajo
        layout.addSpacerItem(QSpacerItem(20, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Layout para botones Aceptar y Cancelar
        button_layout = QHBoxLayout()
        accept_button = QPushButton('Aceptar')
        accept_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        accept_button.clicked.connect(self.save_configuration)
        button_layout.addWidget(accept_button)

        cancel_button = QPushButton('Cancelar')
        cancel_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        # Espaciador para empujar el aviso hacia abajo
        layout.addSpacerItem(QSpacerItem(30, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Agregar el aviso debajo de los botones
        aviso_label = QLabel(
            "Si realiza una modificación en las rutas, debe cerrar la aplicación y volver a abrirla para que todo funcione correctamente."
        )
        aviso_label.setStyleSheet("color: darkorange;")
        aviso_label.setWordWrap(True)
        aviso_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(aviso_label)

    def add_path_layout(self, layout, label_text, config_key, info_callback):
        """Agrega una sección de ruta de archivo al layout."""
        path_layout = QHBoxLayout()
        path_input = QLineEdit(self.config.get(config_key, ''))
        path_button = QPushButton('Buscar')
        path_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        path_button.clicked.connect(lambda: self.open_file_dialog(path_input) if 'Path' in config_key else self.open_directory_dialog(path_input))

        info_button = QPushButton('Inf.')
        info_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        info_button.clicked.connect(info_callback)

        path_layout.addWidget(QLabel(label_text))
        path_layout.addWidget(path_input)
        path_layout.addWidget(info_button)
        path_layout.addWidget(path_button)
        layout.addLayout(path_layout)

    def load_configuration(self):
        if os.path.exists('config.json'):
            with open('config.json', 'r') as config_file:
                return json.load(config_file)
        return {}

    def open_file_dialog(self, line_edit):
        path, _ = QFileDialog.getOpenFileName(self, 'Seleccionar Archivo', '', 'All Files (*)')
        if path:
            line_edit.setText(path)

    def open_directory_dialog(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, 'Seleccionar Directorio')
        if path:
            line_edit.setText(path)

    def save_configuration(self):
        config = {
            'WEBDRIVER_PATH': self.findChild(QLineEdit, 'WebDriver Path:').text(),
            'EDGE_BINARY_PATH': self.findChild(QLineEdit, 'Edge Binary Path:').text(),
            'USER_PROFILE_PATH': self.findChild(QLineEdit, 'User Profile Path:').text()
        }
        with open('config.json', 'w') as config_file:
            json.dump(config, config_file)
        self.accept()

    def show_info_wp(self):
        self.show_info_dialog(
            "Información WEBDRIVER_PATH",
            "<b>Debes especificar la ruta completa al ejecutable del WebDriver</b><br><br>"
            "por ejemplo, <code>C:\\path\\to\\msedgedriver.exe</code>.<br><br>"
            "Este archivo es necesario para que Selenium pueda controlar el navegador Edge."
        )

    def show_info_ep(self):
        self.show_info_dialog(
            "Información EDGE_BINARY_PATH",
            "<b>Debes especificar la ruta completa al ejecutable del navegador Microsoft Edge</b><br><br>"
            "por ejemplo, <code>C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe</code>.<br><br>"
            "Este archivo es necesario para que Selenium pueda iniciar y controlar el navegador."
        )

    def show_info_up(self):
        self.show_info_dialog(
            "Información USER_PROFILE_PATH",
            "<b>Debes especificar la ruta al directorio del perfil de usuario</b><br>"
            "que el navegador debe usar, por ejemplo:<br><br>"
            "<code>C:\\Users\\tu_usuario\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default</code>.<br><br>"
            "<b>Recomendación:</b><br>"
            "Se recomienda abrir ingresado en el navegador al menos una vez para generar el perfil 'Default'."
        )

    def show_info_dialog(self, title, message):
        dialog = InfoDialog(title, message, self)
        dialog.exec_()

class InfoDialog(QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon('D:\\Comercial - Daniel\\BOT_MFS_PAF\\Static\\information icon.webp'))  # Reemplaza 'ruta/a/tu/icono.png' con la ruta a tu archivo de icono

        layout = QVBoxLayout(self)

        self.label = QLabel(message)
        self.label.setTextFormat(Qt.RichText)  # Habilitar el formato Rich Text para QLabel
        layout.addWidget(self.label)

        # Crear un QHBoxLayout para los botones
        button_layout = QHBoxLayout()

        self.okButton = QPushButton('OK')
        self.okButton.clicked.connect(self.accept)
        button_layout.addWidget(self.okButton, alignment=Qt.AlignCenter)

        # Agregar el layout de los botones al layout principal
        layout.addLayout(button_layout)

# Uso de las funciones de información en el contexto de una ventana principal
if __name__ == "__main__":
    app = QApplication([])
    window = ConfigurationDialog()
    window.show()
    app.exec_()
