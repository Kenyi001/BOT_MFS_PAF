import os
import logging
from werkzeug.utils import secure_filename
from flask import Flask, request, jsonify, render_template, send_from_directory, url_for, flash
from selenium_scripts.new_PAF_PTM import ejecutar_automatizacion_new
from selenium_scripts.extraccion_datos_web import extraer_datos_web
from selenium_scripts.Modification_PAF_PTM import ejecutar_automatizacion_modification_new
from config import WEBDRIVER_PATH, EDGE_BINARY_PATH, USER_PROFILE_PATH, URL_INICIO_SESION
app = Flask(__name__)

download_url = None

UPLOAD_FOLDER = r'C:\Users\Kenji\Documents\GitHub\BOT_MFS_PAF\data\Uploaded_Files'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

# Configura el logging
logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        files = request.files.getlist('files[]')

        if not files or all(file.filename == '' for file in files):
            return jsonify({'error': 'No files uploaded'}), 400

        username = request.form.get('username')
        password = request.form.get('password')
        action = request.form.get('action')

        # Captura las selecciones de los checkboxes
        opciones_seleccionadas = {
            'modificar_nombre_responsable': 'nombre_responsable' in request.form,
            'modificar_direccion': 'direccion' in request.form,
            # Añade aquí el resto de las opciones según corresponda
        }

        if not all([username, password, action]):
            return jsonify({'error': 'Missing form data'}), 400

        errors = []
        for file in files:
            if not allowed_file(file.filename):
                errors.append(f'Invalid file type for {file.filename}')
                continue

            try:
                if action == 'extraccionDatosWeb':
                    filename = extraer_datos_web(URL_INICIO_SESION, file, username, password)
                    download_url = url_for('download_file', filename=filename)
                elif action == 'newPAFPTM':
                    response = ejecutar_automatizacion_new(URL_INICIO_SESION, file, username, password)
                    if response.status_code != 200:
                        return jsonify({'error': 'An error occurred', 'details': response.get_data(as_text=True)}), response.status_code
                elif action == 'modificacionPAFPTM':
                    # Pasa las opciones seleccionadas a la función
                    ejecutar_automatizacion_modification_new(file, username, password, opciones_seleccionadas)

                filename = secure_filename(file.filename)
                save_processed_file(file)  # Guarda el archivo procesado después de procesarlo
                download_url = url_for('download_file', filename=filename)

            except Exception as e:
                error_message = f'Error processing {file.filename}: {str(e)}'
                errors.append(error_message)
                logging.error(error_message)
                return jsonify({'error': 'An error occurred while processing the files', 'details': errors}), 502

        return jsonify({'message': 'All files processed successfully', 'download_url': download_url}), 200
    except Exception as e:
        return jsonify({'error': 'An error occurred while processing the files', 'details': str(e)}), 502
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['xlsx']

def save_processed_file(file):
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)  # Guarda el archivo procesado en el directorio UPLOAD_FOLDER
    return file_path

if __name__ == '__main__':
    app.run(host="0.0.0.0",port=4000, debug=True)