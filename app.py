from flask import Flask, render_template, request, redirect, flash
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

EMAIL_ADDRESS = 'tesoreria@dimensasl.com'
EMAIL_PASSWORD = 'Ma.3618d.'

SAVE_FOLDER = 'formularios_guardados'
os.makedirs(SAVE_FOLDER, exist_ok=True)

@app.route('/plantas', methods=['GET', 'POST'])
def formulario_plantas():
    if request.method == 'POST':
        plantas_data = request.form.to_dict()

        # Guardar en Excel
        df = pd.DataFrame([plantas_data])
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_path = os.path.join(SAVE_FOLDER, f'alta_planta_{timestamp}.xlsx')
        df.to_excel(file_path, index=False)

        # Enviar correo
        enviar_correo_aviso(file_path)

        flash('Formulario de planta enviado correctamente.')
        return redirect('/plantas')

    # Si es GET, mostrar formulario
    return render_template('plantas.html')

def enviar_correo_aviso(file_path):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_ADDRESS
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
        print(f'Error adjuntando el archivo: {e}')

    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        print('Correo enviado correctamente.')
    except Exception as e:
        print(f'Error enviando el correo: {e}')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
