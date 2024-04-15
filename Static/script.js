src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"
src="https://cdn.jsdelivr.net/npm/sweetalert2@10"
function handleError(xhr, status, error) {
    var message = "Error desconocido. Por favor intenta de nuevo.";
    switch (xhr.status) {
        case 400:
            message = "Datos inválidos. Por favor revisa tu entrada y asegúrate de que todos los campos están completos.";
            break;
        case 401:
            message = "Autenticación fallida. Verifica tu usuario y contraseña.";
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
    alert(message);
    updateStatus('red', message);
}

function updateStatus(color, text) {
    var statusIndicator = document.getElementById('statusIndicator');
    statusIndicator.style.backgroundColor = color;
    document.getElementById('statusText').innerText = text;
}

$(document).ready(function() {
    $('#uploadForm').submit(function(e) {
        e.preventDefault();
        submitForm();
    });

    function submitForm() {
        var formData = new FormData($('#uploadForm')[0]);
        updateStatus('green', 'Proceso iniciado...');

        $.ajax({
            url: '/upload',
            type: 'POST',
            data: formData,
            contentType: false,
            processData: false,
            beforeSend: function() {
                updateStatus('yellow', 'En proceso...');
            },
            success: function(data) {
                alert('Archivos cargados con éxito: ' + data.message);
                updateStatus('green', 'Completado con éxito');
                var downloadLink = $('<a>').attr('href', data.download_url).text('Descargar Archivo Procesado');
                $('#downloadLinkContainer').empty().append(downloadLink);
            },
            error: function(xhr, status, error) {
                handleError(xhr, status, error);
            }
        });
    }
});