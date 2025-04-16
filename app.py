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

# Configuración de email desde variables de entorno
EMAIL_ADDRESS = os.environ.get('EMAIL_USER', 'migueladr191194@gmail.com')  # Por si no se define
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS', 'zvup wjjv bwas tebs')  # ⚠️ Usar en entorno seguro

# Ruta donde se guardarán los Excel generados
SAVE_FOLDER = 'formularios_guardados_plantas'
os.makedirs(SAVE_FOLDER, exist_ok=True)


@app.route('/formulario_plantas', methods=['GET'])
def formulario_plantas():
    return render_template('formulario_plantas.html')


@app.route('/guardar_plantas', methods=['POST'])
def guardar_plantas():
    plantas_data = request.form.to_dict()

    # Filtrar campos vacíos
    plantas_filtradas = {k: v for k, v in plantas_data.items() if v}

    # Nombre del archivo con nombre empresa si existe
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_empresa = plantas_data.get('nombre_empresa', 'planta').replace(" ", "_")
    file_path = os.path.join(SAVE_FOLDER, f'{nombre_empresa}_{timestamp}.xlsx')

    # Guardar en Excel
    df = pd.DataFrame([plantas_filtradas])
    df.to_excel(file_path, index=False)

    # Enviar email
    enviar_correo_aviso_plantas(file_path, plantas_data.get('correo_comercial'))

    flash('Formulario de plantas enviado correctamente.')
    return redirect('/formulario_plantas')


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

    # Adjuntar archivo
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

    # Enviar email
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


