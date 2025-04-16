from flask import Flask, render_template, request, redirect, flash, send_file
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os
import logging

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configuración de logs
logging.basicConfig(level=logging.INFO)

# Configuración de email desde variables de entorno (o valores por defecto)
EMAIL_ADDRESS = os.environ.get('EMAIL_USER', 'migueladr191194@gmail.com')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS', 'zvup wjjv bwas tebs')  # ⚠️ Reemplaza por seguridad en producción

# Carpeta donde se guardan los archivos Excel
SAVE_FOLDER = 'formularios_guardados_plantas'
os.makedirs(SAVE_FOLDER, exist_ok=True)

# Rutas que muestran el formulario
@app.route('/')
@app.route('/plantas')
@app.route('/formulario_plantas')
def formulario_plantas():
    return render_template('plantas.html')


@app.route('/guardar_plantas', methods=['POST'])
def guardar_plantas():
    plantas_data = request.form.to_dict()

    # Filtrar datos no vacíos
    plantas_filtradas = {k: v for k, v in plantas_data.items() if v}

    # Crear nombre del archivo con nombre de empresa si existe
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_empresa = plantas_data.get('nombre_empresa', 'planta').replace(" ", "_")
    file_path = os.path.join(SAVE_FOLDER, f'{nombre_empresa}_{timestamp}.xlsx')

    # Guardar Excel
    df = pd.DataFrame([plantas_filtradas])
    df.to_excel(file_path, index=False)

    # Enviar correo
    enviar_correo_aviso_plantas(file_path, plantas_data.get('correo_comercial'))

    flash('Formulario de plantas enviado correctamente.')
    return redirect('/plantas')


def enviar_correo_aviso_plantas(file_path, comercial_email=None):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    destinatarios = ['tesoreria@dimensasl.com']
    if comercial_email:
        destinatarios.append(comercial_email)
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = 'Nuevo formulario de alta de planta recibido'

    body = 'Se ha recibido un nuevo formulario de alta de planta. Se adjunta el archivo Excel.'
    msg.attach(MIMEText(body, 'plain'))

    try:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
            msg.attach(part)
    except Exception as e:
        logging.error(f'Error adjuntando el archivo: {e}')
        return

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        logging.info(f'Correo enviado correctamente a: {", ".join(destinatarios)}')
    except Exception as e:
        logging.error(f'Error enviando el correo: {e}')


@app.route('/descargar-ultimo-planta')
def descargar_ultimo_excel_planta():
    archivos = [f for f in os.listdir(SAVE_FOLDER) if f.endswith('.xlsx')]
    if not archivos:
        return "No hay archivos de plantas para descargar."

    archivos.sort(reverse=True)
    archivo_mas_reciente = archivos[0]
    ruta_completa = os.path.join(SAVE_FOLDER, archivo_mas_reciente)
    return send_file(ruta_completa, as_attachment=True)


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)



