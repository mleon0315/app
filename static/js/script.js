
document.getElementById('startDownload').addEventListener('click', async function (event) {
    event.preventDefault();  // Evita que el formulario se envíe inmediatamente

    var archivo = document.getElementById('archivo').value;
    var estado = document.getElementById('estado').value;
    var carpeta_descargas = document.getElementById('carpeta_descargas').value;

    if (!archivo) {
        document.getElementById('invalidFile').innerHTML = `
                    <p>Por favor, carga un archivo Excel.</p>
                `;
        setTimeout(function () {
            document.getElementById('invalidFile').innerHTML = ''
        }, 5000);
    } if (!estado) {
        document.getElementById('invalidState').innerHTML = `
                    <p>Por favor, selecciona un estado.</p>
                `;
        setTimeout(function () {
            document.getElementById('invalidState').innerHTML = ''
        }, 5000);
    } if (!carpeta_descargas) {
        document.getElementById('invalidWay').innerHTML = `
                    <p>Por favor, ingresa la ruta de descarga.</p>
                `;
        setTimeout(function () {
            document.getElementById('invalidWay').innerHTML = ''
        }, 5000);
        return;
    }

    var form = document.getElementById('downloadForm');
    var formData = new FormData(form);

    try {
        let response = await fetch('/', {
            method: 'POST',
            body: formData
        });

        let data = await response.json();

        if (Array.isArray(data)) {
            for (let item of data) {
                await Swal.fire({
                    icon: item.tipo,
                    text: item.mensaje,
                });    
            }
            
        } else {
            Swal.fire({
                icon: data.tipo || 'info',
                text: data.mensaje || 'No hay mensajes para mostrar',
            });
        }
        // Después de mostrar los mensajes, redirigir
        setTimeout(() => {
            window.location.href = '/';  // Redirigir a la URL deseada
        }, 3000);  // Esperar 3 segundos antes de redirigir

    } catch (error) {
        Swal.fire({
            icon: 'error',
            text: 'Ocurrió un error al procesar la solicitud: ' + error.message,
        });
    }; 

// VERIFICA EL PROGRESO PERIODICAMENTE
var progressInterval = setInterval(function () {
    fetch('/get_progress')
        .then(response => response.json())
        .then(data => {
            var progress = data.progress;
            var progressBar = document.getElementById('progressBar');
            progressBar.style.width = progress + '%';
            progressBar.setAttribute('aria-valuenow', progress);
            progressBar.textContent = progress + '%';

            if (progress >= 100) {
                clearInterval(progressInterval);
            }
        });
}, 1000);  // VERIFICA EL PROGRESO CADA SEGUNDO

});
