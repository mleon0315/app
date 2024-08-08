from flask import Flask, render_template, request, jsonify
import os, time, openpyxl, secrets
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from flask_sqlalchemy import SQLAlchemy

# GENERAR CLAVE SECRETA ALEATORIA
secret_key = secrets.token_hex(16)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///site.db'
app.config['SECRET_KEY'] = secret_key
db = SQLAlchemy(app)

# MODELO PARA ALMACENAR EL ARCHIVO EN LA BASE DE DATOS
class ExcelFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(120), nullable=False)
    filepath = db.Column(db.String(120), nullable=False)

# LISTA PARA GUARDAR LOS MENSAJES
mensajes = []

# FUNCION PARA AGREGAR MENSAJES A LA LISTA
def agregar_mensaje(tipo, mensaje):
    mensajes.append({'tipo': tipo, 'mensaje': mensaje})

# FUNCION PARA LEER LOS CUFES DESDE UN ARCHIVO EN EXCEL
def leer_cufes_desde_excel(archivo_excel, estado_seleccionado):
    libro = openpyxl.load_workbook(archivo_excel)
    hoja = libro.active
    cufes = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        cufe = fila[1]
        estado = fila[13]
        if estado == estado_seleccionado:
            cufes.append(cufe)
    libro.close() # CIERRA EL ARCHIVO
    return cufes

# FUNCION PARA BUSCAR Y DESCARGAR UNA FACTURA USANDO EL CUFE
def buscar_y_descargar_factura(driver, cufe, carpeta_descargas):
    nombre_archivo = os.path.join(carpeta_descargas, f'{cufe}.pdf')
    if os.path.exists(nombre_archivo):
        agregar_mensaje('info', f'La factura para el CUFE {cufe} ya existe. No se descargará nuevamente.')
        return

    intentos = 0
    max_intentos = 10

    while intentos < max_intentos:
        driver.get('https://catalogo-vpfe.dian.gov.co/User/SearchDocument')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'DocumentKey')))

        # EJECUTAR EL SCRIPT DE reCAPTCHA
        driver.execute_script("""
            grecaptcha.ready(function() {
                grecaptcha.execute('6LcFcqEUAAAAAMoQN_j0g7tiTS8IQbHcRXfWcyYh', {action: 'adminLogin'}).then(function(token) {
                    document.querySelector(".RecaptchaToken").value = token;
                });
            });
        """)

        # ESPERAR A QUE EL TOKEN SE GENERE
        time.sleep(3)  # Opcional: puedes ajustar este tiempo según tus necesidades

        # INGRESAR EL CUFE Y ENVIAR EL FORMULARIO
        try:
            campo_cufe = driver.find_element(By.ID, 'DocumentKey')
            campo_cufe.send_keys(cufe)
            campo_cufe.send_keys(Keys.RETURN)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="html-gdoc"]/div[3]/div/div[1]/div[3]/p/a'))
            )
            enlace_descarga = driver.find_element(By.XPATH, '//*[@id="html-gdoc"]/div[3]/div/div[1]/div[3]/p/a')
            enlace_descarga.click()
            time.sleep(5)  # ESPERAR A QUE SE COMPLETE LA DESCARGA
            break
        except Exception as e:
            intentos += 1
            agregar_mensaje('error', f'Error al intentar descargar la factura para CUFE {cufe}: {e}')
            if intentos < max_intentos:
                agregar_mensaje('info', f'Reintentando en 6 segundos... (Intento {intentos}/{max_intentos})')
                time.sleep(6)
            else:
                agregar_mensaje('warning', f'Se alcanzó el máximo de {max_intentos} intentos para el CUFE {cufe}. Pasando al siguiente.')

# CONFIGURAR DESCARGAS
def configurar_descargas(carpeta_descargas):
    if not os.path.exists(carpeta_descargas):
        os.makedirs(carpeta_descargas)
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": carpeta_descargas, "download.prompt_for_download": False, "directory_upgrade": True}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return driver

# RUTA PRINCIPAL PARA LA CARGA DEL ARCHIVO Y SELECCION DE LA CARPETA DESCARGAS
@app.route('/', methods=['GET', 'POST'])
def index():
    global mensajes
    mensajes = []
    
    if request.method == 'POST':
        archivo_excel = request.files['archivo']
        carpeta_descargas = request.form['carpeta_descargas']
        estado_seleccionado = request.form['estado']  # Obtener el estado seleccionado
        archivo_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], archivo_excel.filename)
        archivo_excel.save(archivo_excel_path)

        # GUARDAR EN LA BASE DE DATOS
        excel_file = ExcelFile(filename=archivo_excel.filename, filepath=archivo_excel_path)
        db.session.add(excel_file)
        db.session.commit()

        cufes = leer_cufes_desde_excel(archivo_excel_path, estado_seleccionado)
        driver = configurar_descargas(carpeta_descargas)

        total_cufes = len(cufes)
        for index, cufe in enumerate(cufes):
            buscar_y_descargar_factura(driver, cufe, carpeta_descargas)
            #flash(f'Descargando factura {index + 1} de {len(cufes)}...')
            #Actualizar progreso
            with open ('progress.txt', 'w') as f:
                f.write(str(int((index + 1) / total_cufes * 100)))

            time.sleep(2)  # SIMULAR EL PROGRESO

        driver.quit()

        # DESPUES DE COMPLETAR LA DESCARGA
        try:
            # ESPERAR PARA ASEGURAR QUE TODOS LOS PROCESOS HAYAN TERMINADO
            time.sleep(1)
            
    
            # ELIMINAR EL ARCHIVO DE LA CARPETA UPLOADS
            if os.path.exists(archivo_excel_path):
                os.remove(archivo_excel_path)
                print(f"Archivo eliminado de la carpeta: {archivo_excel_path}")
    
            # ELIMINAR EL ARCHIVO DE LA BASE DE DATOS
            db.session.delete(excel_file)
            db.session.commit()
            print(f'Archivo eliminado de la base de datos: {archivo_excel.filename}')
            
            # ELIMINAR EL ARCHIVO DE PROGRESO
            if os.path.exists('progress.txt'):
                os.remove('progress.txt')

        except PermissionError as e:
            print(f"No se pudo eliminar el archivo: {e}")
        except Exception as e:
            print(f"Error desconocido al intentar eliminar el archivo: {e}")

        # MENSAJE
        agregar_mensaje('success','Descargas completadas con éxito')
        
        return jsonify(mensajes)    
        #return redirect(url_for('index'))
        
    return render_template('index.html')

@app.route('/get_progress')
def get_progress():
    if os.path.exists('progress.txt'):
        with open('progress.txt', 'r') as f:
            progress = f.read()
    else:
        progress = '0'
    return jsonify({'progress': int(progress)})

if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    with app.app_context():
        # CREAR LA BASE DE DATOS
        db.create_all()

    app.run(debug=True)
