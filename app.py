from flask import Flask, render_template, request, redirect, flash, send_file
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from openpyxl import load_workbook
import os
import io
import logging

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configuración de logs
logging.basicConfig(level=logging.INFO)

# Email
EMAIL_ADDRESS = os.environ.get('EMAIL_USER', 'migueladr191194@gmail.com')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS', 'zvup wjjv bwas tebs')  # ⚠️ Cambiar en producción

# Carpeta para guardar Excel generados
SAVE_FOLDER = 'formularios_guardados_plantas'
os.makedirs(SAVE_FOLDER, exist_ok=True)

# Mostrar formulario de alta de plantas
@app.route('/')
@app.route('/plantas')
@app.route('/formulario_plantas')
def formulario_plantas():
    return render_template('plantas.html')

# Guardar formulario y enviar Excel
@app.route('/guardar_plantas', methods=['POST'])
def guardar_plantas():
    data = request.form.to_dict()

    # Crear Excel desde plantilla
    excel_mem = crear_excel_plantas_solas(data)

    # Guardar archivo temporal
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_cliente = data.get('nombre_cliente', 'cliente').replace(" ", "_")
    filename = f'{nombre_cliente}_{timestamp}.xlsx'
    file_path = os.path.join(SAVE_FOLDER, filename)
    with open(file_path, 'wb') as f:
        f.write(excel_mem.read())

    # Enviar por correo
    enviar_correo_aviso_plantas(file_path, data.get('correo_comercial'))

    # Mostrar pantalla de agradecimiento
    return render_template('gracias.html')

# Función para generar el Excel con datos de plantas (orden corregido)
def crear_excel_plantas_solas(data):
    wb = load_workbook("Copia de alta de plantas solas.xlsx")
    ws = wb.active

    # A2: Nombre del cliente
    ws["A2"] = data.get("nombre_cliente", "")

    columnas = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
    campos = [
        "planta_nombre_{}", "planta_direccion_{}", "planta_cp_{}", "planta_poblacion_{}",
        "planta_provincia_{}", "planta_telefono_{}", "planta_email_{}", "planta_horario_{}",
        "planta_observaciones_{}", "planta_contacto_nombre_{}", "planta_contacto_telefono_{}",
        "planta_contacto_email_{}"
    ]

    for i in range(1, 11):
        fila = 3 + i  # Desde fila 4 en la plantilla
        valores = [data.get(campo.format(i), "") for campo in campos]
        if not valores[0]:  # Si no hay nombre de planta, se omite
            continue
        for col, val in zip(columnas, valores):
            ws[f"{col}{fila}"] = val

    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    return mem

# Enviar email con el Excel adjunto
def enviar_correo_aviso_plantas(file_path, comercial_email=None):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    destinatarios = ['tesoreria@dimensasl.com']
    if comercial_email:
        destinatarios.append(comercial_email)
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = 'Nuevo formulario de alta de plantas (solo plantas)'

    body = 'Se ha recibido un formulario con alta de plantas. Se adjunta el archivo Excel.'
    msg.attach(MIMEText(body, 'plain'))

    try:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)
    except Exception as e:
        logging.error(f'❌ Error adjuntando archivo: {e}')
        return

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        logging.info('✅ Correo enviado correctamente')
    except Exception as e:
        logging.error(f'❌ Error enviando el correo: {e}')

# Descargar último archivo generado
@app.route('/descargar-ultimo-planta')
def descargar_ultimo_excel_planta():
    archivos = [f for f in os.listdir(SAVE_FOLDER) if f.endswith('.xlsx')]
    if not archivos:
        return "No hay archivos de plantas para descargar."

    archivos.sort(reverse=True)
    ruta_completa = os.path.join(SAVE_FOLDER, archivos[0])
    return send_file(ruta_completa, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)


