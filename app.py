import os
from werkzeug.utils import secure_filename
from flask import Flask, request, jsonify, render_template, send_from_directory, url_for
from selenium_scripts.new_PAF_PTM import ejecutar_automatizacion_new
from selenium_scripts.extraccion_datos_web import extraer_datos_web
from selenium_scripts.Modification_PAF_PTM import ejecutar_automatizacion_modification_new

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'data')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_file():
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
                extraer_datos_web("https://appweb.asfi.gob.bo/RMI/Default.aspx", file, username, password)
            elif action == 'newPAFPTM':
                ejecutar_automatizacion_new(file, username, password)
            elif action == 'modificacionPAFPTM':
                # Pasa las opciones seleccionadas a la función
                ejecutar_automatizacion_modification_new(file, username, password, opciones_seleccionadas)

            filename = secure_filename(file.filename)
            save_processed_file(file)  # Guarda el archivo procesado después de procesarlo
            download_url = url_for('download_file', filename=filename)

        except Exception as e:
            errors.append(f'Error processing {file.filename}: {str(e)}')

    if errors:
        return jsonify({'error': 'Some files were not processed successfully', 'details': errors}), 400

    return jsonify({'message': 'All files processed successfully', 'download_url': download_url})

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['xlsx']

def save_processed_file(file):
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)  # Guarda el archivo procesado en el directorio UPLOAD_FOLDER
    return file_path

if __name__ == '__main__':
    app.run(debug=True)