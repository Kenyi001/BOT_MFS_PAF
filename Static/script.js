src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"
src="https://cdn.jsdelivr.net/npm/sweetalert2@10"
function handleError(xhr, status, error, mefError) {
    var title = "Error";  // Define la variable title
    var message = "Error desconocido. Por favor intenta de nuevo.";
    switch (xhr.status) {
        case 501:
            message = "Error en el archivo. Por favor verifica que el archivo sea un archivo de Excel válido y que contenga la información correcta";
            break;
        case 502:
            message = "Error en el registro con MEF " + mefError + ". Por favor verifique el excel.";
            break;
        case 400:
            message = "Datos inválidos. Por favor revisa tu entrada y asegúrate de que todos los campos están completos.";
            break;
        case 401:
            message = "Autenticación fallida. Verifica tu usuario y contraseña.";
            break;
        case 702:
            message = "No se ha seleccionado ningún archivo. Por favor selecciona un archivo para subir.";
            break;
        case 404:
            message = "El servidor no pudo encontrar lo solicitado. Por favor intenta de nuevo más tarde.";
            break;
        case 413:
            message = "El archivo subido es demasiado grande. Por favor verifica los límites de tamaño de archivo.";
            break;
        case 500:
            message = "Problema interno del servidor. Estamos trabajando para resolverlo lo antes posible.";
            break;
        case 0:
            message = "Problema de conexión. Verifica tu conexión a Internet.";
            break;
    }
    updateStatus('red', message);
    setTimeout(() => {
        showMessage(title, message, 'error');
    }, 100);
}

function updateStatus(color, text) {
    var statusIndicator = document.getElementById('statusIndicator');
    statusIndicator.style.backgroundColor = color;
    document.getElementById('statusText').innerText = text;
}

function showMessage(title, message, type) {
    // Aquí puedes agregar lógica para manejar diferentes tipos de mensajes
    // Por ahora, solo vamos a anteponer el tipo al título
    var fullTitle = type.toUpperCase() + ": " + title;

    // Mostrar el mensaje
    alert(fullTitle + "\n\n" + message);
}

$(document).ready(function() {
    $('#uploadForm').submit(function(e) {
        e.preventDefault();
        var confirmation = confirm("¿Estás seguro de que quieres enviar el formulario?");
        if (confirmation) {
            submitForm();
        }
    });

    function submitForm() {
        var formData = new FormData($('#uploadForm')[0]);
        updateStatus('green', 'Proceso iniciado...');
        console.log("Formulario enviado, iniciando solicitud AJAX...");

        $.ajax({
            url: '/upload',
            type: 'POST',
            data: formData,
            contentType: false,
            processData: false,
            beforeSend: () => {
                updateStatus('yellow', 'En proceso...');
                $('#spinner').show();
                console.log("Solicitud AJAX en proceso...");
            },
            success: data => {
                console.log("Solicitud AJAX completada con éxito.");
                updateStatus('green', 'Completado con éxito');
                if (data.download_url) {
                    $('#spinner').attr('href', data.download_url).show();
                }
                setTimeout(() => {
                    showMessage('Éxito', `Archivos cargados con éxito: ${data.message}`, 'success');
                }, 100);
                $('#spinner').hide();
            },
            error: function(xhr, status, error) {
                console.log("Error en la solicitud AJAX: ", error);
                console.log("Respuesta completa del servidor: ", xhr.responseText);
                try {
                    var response = JSON.parse(xhr.responseText);
                    var message = response.error;
                    var details = response.details.join('\n');
                    var mefError = response.mef ? response.mef : "undefined";
                    handleError(xhr, status, error, mefError);
                    $('#spinner').hide();
                    setTimeout(() => {
                        showMessage('Error', `${message}\n\n${details}`, 'error');
                    }, 100);
                } catch(e) {
                    console.log("Error parsing JSON response: ", e);
                }
            }
        });
    }
});
$(window).on('load', function() {
    $('#preloader').fadeOut('slow');
});