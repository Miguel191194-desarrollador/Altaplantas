from flask import Flask, render_template, request, redirect, flash, send_file, url_for
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from openpyxl import load_workbook
import os
import logging
import threading
# Carga variables desde .env si existe (opcional)
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Logs
logging.basicConfig(level=logging.INFO)

# === Configuración de email (usar variables de entorno) ===
# ATENCIÓN: Deben existir EMAIL_USER y EMAIL_PASS en el entorno
EMAIL_ADDRESS = os.environ.get('EMAIL_USER')    # ej. tucuenta@gmail.com
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')   # contraseña de aplicación (16 chars)

# Carpeta para guardar Excel generados
SAVE_FOLDER = 'formularios_guardados_plantas'
os.makedirs(SAVE_FOLDER, exist_ok=True)

# ---------------- Rutas ----------------

@app.route('/')
@app.route('/plantas')
@app.route('/formulario_plantas')
def formulario_plantas():
    return render_template('plantas.html')

@app.route('/guardar_plantas', methods=['POST'])
def guardar_plantas():
    data = request.form.to_dict()

    # 1) Generar y guardar Excel en disco
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_cliente = (data.get('nombre_cliente') or 'cliente').replace(" ", "_")
    filename = f'{nombre_cliente}_{timestamp}.xlsx'
    file_path = os.path.join(SAVE_FOLDER, filename)
    try:
        crear_excel_plantas_solas_a_archivo(data, file_path)
    except Exception as e:
        logging.exception("❌ Error generando el Excel de plantas")
        flash("Ha ocurrido un error generando el Excel.")
        return redirect(url_for('formulario_plantas'))

    # 2) Enviar correo en segundo plano para no bloquear la respuesta
    destinatario_comercial = data.get('correo_comercial')
    threading.Thread(target=enviar_correo_aviso_plantas, args=(file_path, destinatario_comercial), daemon=True).start()

    # 3) Pantalla de agradecimiento
    return render_template('gracias.html')

@app.route('/descargar-ultimo-planta')
def descargar_ultimo_excel_planta():
    archivos = [f for f in os.listdir(SAVE_FOLDER) if f.endswith('.xlsx')]
    if not archivos:
        return "No hay archivos de plantas para descargar."
    archivos.sort(reverse=True)
    ruta_completa = os.path.join(SAVE_FOLDER, archivos[0])
    return send_file(ruta_completa, as_attachment=True)

# (opcional) Ruta para verificar que las variables están cargadas
@app.route("/_env")
def _env():
    ok_user = "OK" if os.getenv("EMAIL_USER") else "MISSING"
    ok_pass = "OK" if os.getenv("EMAIL_PASS") else "MISSING"
    return f"EMAIL_USER: {ok_user} | EMAIL_PASS: {ok_pass}"

# -------------- Lógica de Excel --------------

def crear_excel_plantas_solas_a_archivo(data, file_path):
    """
    Plantilla: 'Copia de alta de plantas solas.xlsx'
    Reglas:
      - A2 = Nombre del cliente
      - Las plantas empiezan en la fila 5
      - Orden de columnas: B, D, C, E, F, G, H, I, J, K, L, M
        (Nombre, Dirección, CP, Población, Provincia, Teléfono, Email,
         Horario, Observaciones, ContactoNombre, ContactoTeléfono, ContactoEmail)
    """
    wb = load_workbook("Copia de alta de plantas solas.xlsx")
    ws = wb.active

    # A2: Nombre del cliente
    ws["A2"] = data.get("nombre_cliente", "")

    # Columnas en el orden correcto para "plantas solas"
    columnas = ["B", "D", "C", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
    campos = [
        "planta_nombre_{}", "planta_direccion_{}", "planta_cp_{}", "planta_poblacion_{}",
        "planta_provincia_{}", "planta_telefono_{}", "planta_email_{}", "planta_horario_{}",
        "planta_observaciones_{}", "planta_contacto_nombre_{}", "planta_contacto_telefono_{}",
        "planta_contacto_email_{}"
    ]

    # Fila de inicio 5 (para i=1 -> fila 5)
    for i in range(1, 11):
        fila = 4 + i
        valores = [data.get(campo.format(i), "") for campo in campos]
        if not (valores[0] or "").strip():  # sin nombre de planta -> omitir
            continue
        for col, val in zip(columnas, valores):
            ws[f"{col}{fila}"] = val

    wb.save(file_path)

# -------------- Envío de correo --------------

def enviar_correo_aviso_plantas(file_path, comercial_email=None):
    """
    Envío por SMTP Gmail (SSL 465). Requiere:
      - EMAIL_USER (remitente Gmail)
      - EMAIL_PASS (contraseña de aplicación)
    """
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        logging.error("❌ Faltan variables EMAIL_USER o EMAIL_PASS. Configúralas antes de enviar.")
        return

    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    destinatarios = ['tesoreria@dimensasl.com']
    if comercial_email and "@" in comercial_email:
        destinatarios.append(comercial_email)
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = 'Nuevo formulario de alta de plantas (solo plantas)'

    body = 'Se ha recibido un formulario con alta de plantas. Se adjunta el archivo Excel.'
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    # Adjuntar Excel
    try:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)
    except Exception:
        logging.exception('❌ Error adjuntando archivo')
        return

    # Enviar
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context, timeout=20) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, destinatarios, msg.as_string())
        logging.info('✅ Correo enviado correctamente')
    except Exception:
        logging.exception('❌ Error enviando el correo (SMTP)')

# -------------- Main --------------

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)


