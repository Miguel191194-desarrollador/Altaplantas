from flask import Flask, render_template, request, redirect, flash, send_file
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configuración de email
EMAIL_ADDRESS = 'migueladr191194@gmail.com'  # Tu correo ✅
EMAIL_PASSWORD = 'zvup wjjv bwas tebs'  # Tu contraseña ✅

# Ruta donde se guardarán los Excel generados
SAVE_FOLDER = 'formularios_guardados_plantas'
os.makedirs(SAVE_FOLDER, exist_ok=True)


@app.route('/formulario_plantas', methods=['GET'])
def formulario_plantas():
    return render_template('formulario_plantas.html')


@app.route('/guardar_plantas', methods=['POST'])
def guardar_plantas():
    plantas_data = request.form.to_dict()

    # Filtramos las plantas que tienen datos
    plantas_filtradas = {}
    for key, value in plantas_data.items():
        if value:  # Si el valor no está vacío
            plantas_filtradas[key] = value

    # Guardar en Excel
    df = pd.DataFrame([plantas_filtradas])
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    file_path = os.path.join(SAVE_FOLDER, f'alta_planta_{timestamp}.xlsx')
    df.to_excel(file_path, index=False)

    # Enviar correo de aviso con adjunto
    enviar_correo_aviso_plantas(file_path, plantas_data.get('correo_comercial'))

    # Mensaje de éxito
    flash('Formulario de plantas enviado correctamente.')
    return redirect('/formulario_plantas')


def enviar_correo_aviso_plantas(file_path, comercial_email=None):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    destinatarios = ['tesoreria@dimensasl.com']
    if comercial_email:
        destinatarios.append(comercial_email)
    msg['To'] = ', '.join(destinatarios)  # Unir destinatarios en una cadena
    msg['Subject'] = 'Nuevo formulario de alta de planta recibido'

    body = 'Se ha recibido un nuevo formulario de alta de planta. Se adjunta el archivo Excel.'
    msg.attach(MIMEText(body, 'plain'))

    # Adjuntar el archivo Excel
    try:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={os.path.basename(file_path)}',
            )
            msg.attach(part)
    except Exception as e:
        print(f'Error adjuntando el archivo: {e}')

    # Enviar el correo
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        print('Correo enviado correctamente a:', ', '.join(destinatarios))
    except Exception as e:
        print(f'Error enviando el correo: {e}')


# ✅ Ruta para descargar el archivo Excel más reciente de plantas
@app.route('/descargar-ultimo-planta')
def descargar_ultimo_excel_planta():
    archivos = [f for f in os.listdir(SAVE_FOLDER) if f.endswith('.xlsx')]
    if not archivos:
        return "No hay archivos de plantas para descargar."

    archivos.sort(reverse=True)  # Ordena por fecha en nombre
    archivo_mas_reciente = archivos[0]
    ruta_completa = os.path.join(SAVE_FOLDER, archivo_mas_reciente)

    return send_file(ruta_completa, as_attachment=True)


if __name__ == '__main__':
    # Cambiado para que funcione en Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

