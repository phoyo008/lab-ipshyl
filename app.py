"""
AUTOMATIZADOR DE ENVÍO DE RESULTADOS DE LABORATORIO
IPS H&L Salud

Flujo principal:
  1. El usuario sube un PDF de resultado de laboratorio (SYNLAB o COLCAN)
  2. Se extrae el nombre, documento, referencia y fecha del PDF
  3. Se busca el email del paciente en Google Sheets por número de documento
  4. Se valida el email y se envía el PDF como adjunto por Gmail
  5. Se marca la columna ENVIADO=Si en el Google Sheet
  6. Se guarda un registro en la base de datos local SQLite
"""

from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.utils import secure_filename
from functools import wraps
import os
import re
import uuid
import sqlite3
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import pdfplumber
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

load_dotenv()

# ==================== CONFIGURACIÓN ====================
# Almacena temporalmente los envíos pendientes de confirmación.
# Clave: token UUID  |  Valor: dict con filepath, datos del paciente y timestamp
pending_sends = {}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'           # Carpeta temporal para PDFs subidos
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Límite de 16 MB por archivo
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-key-change-in-production')

ALLOWED_EXTENSIONS = {'pdf'}

# Credenciales de acceso al sistema web (configuradas en .env)
APP_USERNAME = os.getenv('APP_USERNAME', 'laboratorio')
APP_PASSWORD = os.getenv('APP_PASSWORD', 'ipshyl2025')

# Credenciales de correo (configuradas en .env)
GMAIL_EMAIL    = os.getenv('GMAIL_EMAIL')
GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')

# Servidor SMTP configurable — por defecto Gmail, en Railway usar Brevo u otro
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT   = int(os.getenv('SMTP_PORT', '587'))
SMTP_USER   = os.getenv('SMTP_USER', GMAIL_EMAIL)
SMTP_PASS   = os.getenv('SMTP_PASS', GMAIL_PASSWORD)

# Google Sheets: nombre de la hoja y archivo de credenciales del service account
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'Directorio_IPS')
CREDENTIALS_FILE  = 'clave.json'

# Base de datos SQLite local para el historial de envíos
DB_PATH = 'reportes.db'


# ==================== AUTENTICACIÓN ====================

def login_required(f):
    """Decorador que redirige al login si el usuario no ha iniciado sesión."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username == APP_USERNAME and password == APP_PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        error = 'Usuario o contraseña incorrectos'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ==================== BASE DE DATOS ====================

def init_db():
    """
    Crea la tabla 'envios' si no existe.
    Se llama al iniciar la aplicación.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS envios (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_envio   TEXT NOT NULL,
            nombre        TEXT,
            documento     TEXT,
            email         TEXT,
            referencia    TEXT,
            fecha_resultado TEXT,
            estado        TEXT
        )
    ''')
    conn.commit()
    conn.close()


def save_to_db(nombre, documento, email, referencia, fecha_resultado, estado):
    """
    Inserta un registro en la tabla 'envios'.
    El estado puede ser 'Enviado' o un mensaje de error como 'Error: ...'.
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            '''INSERT INTO envios (fecha_envio, nombre, documento, email, referencia, fecha_resultado, estado)
               VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
             nombre, documento, email, referencia, fecha_resultado, estado)
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error guardando en DB: {e}")


# ==================== FUNCIONES AUXILIARES ====================

def allowed_file(filename):
    """Verifica que el archivo tenga extensión .pdf."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_sheets_client():
    """
    Retorna un cliente de gspread autenticado con el service account.
    Lee las credenciales desde la variable de entorno GOOGLE_CREDENTIALS_JSON
    (recomendado para producción en la nube) o desde el archivo clave.json (desarrollo local).
    """
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if credentials_json:
        import json
        info = json.loads(credentials_json)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)

    return gspread.authorize(creds)


# ==================== EXTRACCIÓN DE DATOS DEL PDF ====================

def extract_pdf_data(filepath):
    """
    Extrae los datos del paciente desde el PDF del resultado de laboratorio.

    Laboratorios soportados:
      - SYNLAB:  NOMBRE / DOCUMENTO: CC. / REFERENCIA / FECHA INGRESO
      - COLCAN:  Nombre / Idenficación: CC / barcode numérico / Fecha toma muestra

    Nota COLCAN: su fuente tipográfica defectuosa codifica la letra 't' como byte nulo
    (\\x00). Se limpia el texto antes de aplicar los regex.

    Retorna un dict con claves: nombre, documento, referencia, fecha, resultado_completo.
    Retorna None si ocurre un error al abrir el PDF.
    """
    try:
        # Extraer todo el texto del PDF página por página
        with pdfplumber.open(filepath) as pdf:
            text = ''.join((page.extract_text() or '') + '\n' for page in pdf.pages)

        # Corrección de encoding defectuoso en PDFs de COLCAN:
        # pdfplumber extrae \x00 donde debería estar la letra 't'
        text = text.replace('\x00', 't')

        # NOMBRE — se detiene antes de Nº, REFERENCIA, DOCUMENTO, Iden o Tel:
        nombre_match = re.search(
            r'(?:NOMBRE|Nombre)\s*:\s*(.+?)(?=\s*(?:N[°º]|REFERENCIA|DOCUMENTO|Iden|Tel:|$))',
            text, re.IGNORECASE
        )

        # DOCUMENTO — "Iden\w+" cubre tanto "Identificación" como "Idenficación" (typo de COLCAN)
        documento_match = re.search(
            r'(?:DOCUMENTO|Iden\w+)\s*:\s*CC\.?\s*(\d{5,12})',
            text, re.IGNORECASE
        )

        # REFERENCIA — SYNLAB usa "REFERENCIA: 123"; COLCAN tiene un código de barras numérico solo en su línea
        referencia_match = re.search(r'REFERENCIA\s*:\s*(\d+)', text, re.IGNORECASE)
        if not referencia_match:
            # Fallback: busca un número de 10-15 dígitos en su propia línea (barcode de COLCAN)
            referencia_match = re.search(r'^\s*(\d{10,15})\s*$', text, re.MULTILINE)

        # FECHA — acepta los formatos de SYNLAB y COLCAN
        fecha_match = re.search(
            r'(?:FECHA INGRESO|Fecha toma muestra|Fecha de recepci[oó]n)\s*:\s*(\d{1,2}[/\.\-]\w+[/\.\-]\d{2,4})',
            text, re.IGNORECASE
        )

        return {
            'nombre':             nombre_match.group(1).strip() if nombre_match else None,
            'documento':          documento_match.group(1) if documento_match else None,
            'referencia':         referencia_match.group(1) if referencia_match else None,
            'fecha':              fecha_match.group(1) if fecha_match else None,
            'resultado_completo': text
        }

    except Exception as e:
        print(f"Error extrayendo PDF: {e}")
        return None


# ==================== GOOGLE SHEETS ====================

def _parse_sheet_date(value):
    """
    Convierte un valor de fecha/timestamp del sheet a un objeto datetime para comparar.
    Soporta los formatos más comunes de Google Forms:
      - 'DD/MM/YYYY HH:MM:SS'  (Marca temporal de Forms)
      - 'DD/MM/YYYY'
      - 'YYYY-MM-DD HH:MM:SS'
      - 'YYYY-MM-DD'
    Retorna datetime o datetime.min si no puede parsear (así queda al fondo del orden).
    """
    value = value.strip()
    for fmt in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return datetime.min


def find_email_in_sheets(documento):
    """
    Busca el email del paciente en Google Sheets usando su número de documento.

    Cuando hay varias filas con el mismo documento (paciente que ha ido varias veces),
    selecciona la fila con la fecha más reciente usando la columna de timestamp
    ('Marca temporal' de Google Forms u otra columna de fecha detectada automáticamente).

    Estructura de la hoja:
      - Fila 1: IDs internos (ignorados)
      - Fila 2: Encabezados de columna
      - Fila 3+: Datos de pacientes

    Columnas detectadas por nombre en encabezado:
      - Fecha/timestamp : columna con 'marca' y 'temporal', o 'fecha'
      - Número de documento: columna con 'número' y 'documento'
      - Email de resultados: columna con 'e-mail' y 'resultado'

    Retorna: (email, número_de_fila) o (None, None) si no se encuentra.
    """
    try:
        client = get_sheets_client()
        sheet  = client.open(GOOGLE_SHEET_NAME).sheet1
        rows   = sheet.get_all_values()

        # La fila 2 (índice 1) contiene los encabezados
        headers = rows[1]

        doc_col   = next((i for i, h in enumerate(headers) if 'número' in h.lower() and 'documento' in h.lower()), None)
        email_col = next((i for i, h in enumerate(headers) if 'e-mail' in h.lower() and 'resultado' in h.lower()), None)

        # Columna de fecha: preferir 'Marca temporal' (Google Forms), si no cualquier columna con 'fecha'
        fecha_col = next((i for i, h in enumerate(headers) if 'marca' in h.lower() and 'temporal' in h.lower()), None)
        if fecha_col is None:
            fecha_col = next((i for i, h in enumerate(headers) if 'fecha' in h.lower()), None)

        if doc_col is None or email_col is None:
            print("Columnas no encontradas:", headers)
            return None, None

        # Recopilar TODAS las filas que coincidan con el documento
        coincidencias = []
        for idx, row in enumerate(rows[2:]):
            if len(row) > doc_col and str(row[doc_col]).strip() == str(documento).strip():
                sheet_row = idx + 3  # Número de fila real (1-based) en el sheet
                email     = row[email_col].strip() if len(row) > email_col else ''

                # Obtener el valor de fecha para ordenar
                fecha_val = ''
                if fecha_col is not None and len(row) > fecha_col:
                    fecha_val = row[fecha_col]

                coincidencias.append({
                    'email':      email,
                    'sheet_row':  sheet_row,
                    'fecha_dt':   _parse_sheet_date(fecha_val),
                    'fecha_raw':  fecha_val
                })

        if not coincidencias:
            return None, None

        if len(coincidencias) > 1:
            print(f"Documento {documento}: {len(coincidencias)} registros encontrados.")
            for c in coincidencias:
                print(f"  → fila {c['sheet_row']} | fecha: {c['fecha_raw']} | email: '{c['email']}'")

        # Preferir filas que tengan email; entre esas, tomar la más reciente
        con_email = [c for c in coincidencias if c['email']]

        if con_email:
            mejor = max(con_email, key=lambda x: x['fecha_dt'])
        else:
            # Ninguna fila tiene email → retornar None
            print(f"Documento {documento}: ningún registro tiene email.")
            return None, None

        print(f"  → Seleccionada fila {mejor['sheet_row']} (fecha: {mejor['fecha_raw']}, email: {mejor['email']})")
        return mejor['email'], mejor['sheet_row']

    except Exception as e:
        print(f"Error buscando en Google Sheets: {e}")
        return None, None


def mark_sent_in_sheet(sheet_row):
    """
    Actualiza la columna 'ENVIADO' a 'Si' en la fila indicada del Google Sheet.
    La columna se detecta automáticamente por su encabezado.
    """
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        headers = sheet.row_values(2)  # Encabezados en fila 2

        # Buscar índice de la columna ENVIADO (1-based para update_cell)
        enviado_col = next((i + 1 for i, h in enumerate(headers) if h.strip().upper() == 'ENVIADO'), None)

        if enviado_col is None:
            print("Columna ENVIADO no encontrada")
            return

        sheet.update_cell(sheet_row, enviado_col, 'Si')

    except Exception as e:
        print(f"Error marcando enviado en sheet: {e}")


# ==================== ENVÍO DE EMAIL ====================

def send_email(to_email, patient_name, referencia, fecha, pdf_path):
    """
    Envía un correo adjuntando el PDF del resultado.

    Usa SendGrid API si la variable SENDGRID_API_KEY está configurada (producción en Railway).
    Cae a Gmail SMTP si no está configurada (desarrollo local con App Password).
    Retorna True si el envío fue exitoso, False en caso contrario.
    """
    sendgrid_key = os.getenv('SENDGRID_API_KEY')

    cuerpo = (
        f"Estimado(a) {patient_name.strip()},\n\n"
        f"Adjuntamos sus resultados de laboratorio.\n\n"
        f"REFERENCIA: {referencia}\n"
        f"FECHA:      {fecha}\n\n"
        f"Si tiene preguntas sobre sus resultados, por favor contacte a su médico.\n\n"
        f"Saludos,\n"
        f"IPS H&L Salud - Laboratorio"
    )

    if sendgrid_key:
        # ---- SendGrid (producción en Railway) ----
        try:
            import base64
            from sendgrid import SendGridAPIClient
            from sendgrid.helpers.mail import (
                Mail, Attachment, FileContent, FileName, FileType, Disposition
            )

            mail = Mail(
                from_email=GMAIL_EMAIL,
                to_emails=to_email,
                subject=f'Resultados de laboratorio - Ref: {referencia}',
                plain_text_content=cuerpo
            )

            with open(pdf_path, 'rb') as f:
                pdf_data = base64.b64encode(f.read()).decode()

            attachment = Attachment(
                file_content=FileContent(pdf_data),
                file_name=FileName(f'resultado_{referencia}.pdf'),
                file_type=FileType('application/pdf'),
                disposition=Disposition('attachment')
            )
            mail.attachment = attachment

            sg = SendGridAPIClient(sendgrid_key)
            response = sg.send(mail)
            return response.status_code in (200, 202)

        except Exception as e:
            print(f"Error enviando email (SendGrid): {e}")
            return False

    else:
        # ---- SMTP genérico (Gmail local, Brevo en Railway, etc.) ----
        try:
            mensaje = MIMEMultipart()
            mensaje['From']    = GMAIL_EMAIL
            mensaje['To']      = to_email
            mensaje['Subject'] = f'Resultados de laboratorio - Ref: {referencia}'
            mensaje.attach(MIMEText(cuerpo, 'plain'))

            with open(pdf_path, 'rb') as f:
                adjunto = MIMEBase('application', 'octet-stream')
                adjunto.set_payload(f.read())
            encoders.encode_base64(adjunto)
            adjunto.add_header('Content-Disposition', f'attachment; filename="resultado_{referencia}.pdf"')
            mensaje.attach(adjunto)

            servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            servidor.starttls()
            servidor.login(SMTP_USER, SMTP_PASS)
            servidor.send_message(mensaje)
            servidor.quit()
            return True

        except Exception as e:
            print(f"Error enviando email (SMTP): {e}")
            return False


def log_error(patient_name, documento, error_msg):
    """Registra un error en la base de datos local."""
    save_to_db(patient_name, documento, '', 'N/A', 'N/A', f'Error: {error_msg}')


# ==================== RUTAS ====================

@app.route('/')
@login_required
def index():
    """Página principal: formulario de carga de PDF."""
    return render_template('index.html')


@app.route('/reportes')
@login_required
def reportes():
    """
    Página de historial: muestra todos los envíos registrados en SQLite,
    ordenados del más reciente al más antiguo.
    """
    conn   = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    envios = conn.execute('SELECT * FROM envios ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('reportes.html', envios=envios)


def _cleanup_stale_pending(max_minutes=30):
    """Elimina entradas de pending_sends con más de max_minutes de antigüedad."""
    now = datetime.now()
    stale = [
        token for token, entry in pending_sends.items()
        if (now - entry['timestamp']).total_seconds() > max_minutes * 60
    ]
    for token in stale:
        fp = pending_sends[token].get('filepath')
        if fp and os.path.exists(fp):
            os.remove(fp)
        del pending_sends[token]


@app.route('/api/preview-pdf', methods=['POST'])
@login_required
def preview_pdf():
    """
    PASO 1 del flujo de envío.
    Recibe el PDF, extrae los datos del paciente y busca su email en Google Sheets.
    NO envía ningún correo — solo retorna la información para que el usuario confirme.

    Retorna un token que identifica el envío pendiente, junto con los datos del paciente.
    El token debe enviarse a /api/confirm-send para ejecutar el envío real.
    """
    filepath = None
    try:
        # Limpiar tokens viejos en cada nueva solicitud
        _cleanup_stale_pending()

        if 'file' not in request.files:
            return jsonify({'error': 'No se envió archivo'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Archivo no seleccionado'}), 400
        if not allowed_file(file.filename):
            return jsonify({'error': 'Solo se permiten archivos PDF'}), 400

        # Guardar el PDF — se conserva hasta que el usuario confirme o cancele
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        filename = f"{uuid.uuid4().hex}_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Extraer datos del PDF
        patient_data = extract_pdf_data(filepath)
        if not patient_data or not patient_data.get('documento'):
            log_error('Desconocido', 'N/A', 'No se pudo extraer datos del PDF')
            return jsonify({'error': 'No se pudieron extraer datos del PDF'}), 400

        nombre     = patient_data.get('nombre', 'Desconocido')
        documento  = patient_data.get('documento', 'N/A')
        referencia = patient_data.get('referencia', 'N/A')
        fecha      = patient_data.get('fecha', 'N/A')

        # Buscar email en Google Sheets (elige la fila con fecha más reciente)
        patient_email, sheet_row = find_email_in_sheets(documento)
        if not patient_email:
            log_error(nombre, documento, 'Paciente no encontrado o sin email en directorio')
            return jsonify({
                'error': 'Paciente no encontrado en el directorio',
                'nombre': nombre,
                'documento': documento
            }), 404

        # Validar formato del email
        if not re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', patient_email):
            log_error(nombre, documento, f'Email inválido: {patient_email}')
            return jsonify({'error': f'Email inválido: {patient_email}', 'nombre': nombre}), 400

        # Generar token y guardar el estado pendiente en memoria
        token = uuid.uuid4().hex
        pending_sends[token] = {
            'filepath':  filepath,
            'nombre':    nombre,
            'documento': documento,
            'email':     patient_email,
            'referencia': referencia,
            'fecha':     fecha,
            'sheet_row': sheet_row,
            'timestamp': datetime.now()
        }

        # Retornar preview — el PDF NO se borra aquí
        return jsonify({
            'token':     token,
            'nombre':    nombre,
            'documento': documento,
            'email':     patient_email,
            'referencia': referencia,
            'fecha':     fecha
        }), 200

    except Exception as e:
        # Si algo falla antes de guardar el token, limpiar el archivo
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': f'Error: {str(e)}'}), 500


@app.route('/api/confirm-send', methods=['POST'])
@login_required
def confirm_send():
    """
    PASO 2 del flujo de envío.
    Recibe el token generado por /api/preview-pdf y ejecuta el envío real:
      1. Envía el correo con el PDF adjunto
      2. Marca ENVIADO=Si en Google Sheets
      3. Guarda registro en SQLite
      4. Elimina el PDF temporal
    """
    try:
        data  = request.get_json()
        token = data.get('token') if data else None

        if not token or token not in pending_sends:
            return jsonify({'error': 'Token inválido o expirado. Vuelve a subir el PDF.'}), 400

        entry = pending_sends.pop(token)

        filepath   = entry['filepath']
        nombre     = entry['nombre']
        documento  = entry['documento']
        email      = entry['email']
        referencia = entry['referencia']
        fecha      = entry['fecha']
        sheet_row  = entry['sheet_row']

        try:
            # Enviar email con el PDF adjunto
            success = send_email(
                to_email=email,
                patient_name=nombre,
                referencia=referencia,
                fecha=fecha,
                pdf_path=filepath
            )

            if success:
                # Marcar ENVIADO=Si en Google Sheets
                if sheet_row:
                    mark_sent_in_sheet(sheet_row)

                # Guardar registro en SQLite
                save_to_db(nombre, documento, email, referencia, fecha, 'Enviado')

                return jsonify({
                    'success':   True,
                    'message':   'Email enviado exitosamente',
                    'nombre':    nombre,
                    'email':     email,
                    'documento': documento
                }), 200
            else:
                log_error(nombre, documento, 'Error al enviar email')
                return jsonify({'error': 'Error al enviar email. Verifica credenciales de Gmail.'}), 500

        finally:
            # Eliminar el PDF temporal siempre, sin importar el resultado
            if filepath and os.path.exists(filepath):
                os.remove(filepath)

    except Exception as e:
        return jsonify({'error': f'Error: {str(e)}'}), 500


@app.route('/api/cancel-send', methods=['POST'])
@login_required
def cancel_send():
    """Cancela un envío pendiente y elimina el PDF temporal."""
    try:
        data  = request.get_json()
        token = data.get('token') if data else None

        if token and token in pending_sends:
            fp = pending_sends.pop(token).get('filepath')
            if fp and os.path.exists(fp):
                os.remove(fp)

        return jsonify({'ok': True}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/test-connection', methods=['GET'])
@login_required
def test_connection():
    """
    Prueba las conexiones externas:
      - Google Sheets: intenta abrir la hoja y contar registros
      - Gmail SMTP: intenta autenticarse con las credenciales configuradas

    Retorna 200 si todo está OK, 500 si alguna conexión falla.
    """
    results = {}

    # Prueba Google Sheets
    try:
        client = get_sheets_client()
        sheet  = client.open(GOOGLE_SHEET_NAME).sheet1
        rows   = sheet.get_all_values()
        total  = len(rows) - 2 if len(rows) > 2 else 0  # Descontar fila de IDs y encabezados
        results['google_sheets'] = {
            'ok':      True,
            'mensaje': f'Conectado. {total} registros en el directorio.',
            'hoja':    GOOGLE_SHEET_NAME
        }
    except Exception as e:
        results['google_sheets'] = {'ok': False, 'mensaje': str(e)}

    # Prueba de email: SendGrid (producción) o Gmail SMTP (local)
    sendgrid_key = os.getenv('SENDGRID_API_KEY')
    try:
        if sendgrid_key:
            results['email'] = {'ok': True, 'mensaje': f'SendGrid configurado. Enviando desde {GMAIL_EMAIL}'}
        else:
            if not SMTP_USER:
                raise ValueError('Credenciales de correo no configuradas')
            servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10)
            servidor.starttls()
            servidor.login(SMTP_USER, SMTP_PASS)
            servidor.quit()
            results['email'] = {'ok': True, 'mensaje': f'SMTP conectado ({SMTP_SERVER}) como {SMTP_USER}'}
    except Exception as e:
        results['email'] = {'ok': False, 'mensaje': str(e)}

    todo_ok = all(r['ok'] for r in results.values())
    return jsonify({'ok': todo_ok, 'resultados': results}), 200 if todo_ok else 500


@app.route('/api/debug-sheet', methods=['GET'])
@login_required
def debug_sheet():
    """Muestra los encabezados reales del Google Sheet para diagnosticar problemas de columnas."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        rows    = sheet.get_all_values()
        headers = rows[1] if len(rows) > 1 else []
        return jsonify({
            'total_filas': len(rows),
            'encabezados_fila2': {str(i): h for i, h in enumerate(headers)}
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug-search/<documento>', methods=['GET'])
@login_required
def debug_search(documento):
    """Busca un documento directamente en el sheet y muestra qué columnas detecta y qué filas encuentra."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        rows    = sheet.get_all_values()
        headers = rows[1] if len(rows) > 1 else []

        doc_col   = next((i for i, h in enumerate(headers) if 'número' in h.lower() and 'documento' in h.lower()), None)
        email_col = next((i for i, h in enumerate(headers) if 'e-mail' in h.lower() and 'resultado' in h.lower()), None)

        coincidencias = []
        for idx, row in enumerate(rows[2:]):
            if len(row) > (doc_col or 0):
                val = str(row[doc_col]).strip() if doc_col is not None else ''
                if val == str(documento).strip():
                    coincidencias.append({
                        'fila_sheet': idx + 3,
                        'doc_value':  val,
                        'email':      row[email_col].strip() if email_col is not None and len(row) > email_col else 'COL_FUERA_DE_RANGO',
                        'nombre':     row[6].strip() if len(row) > 6 else ''
                    })

        return jsonify({
            'documento_buscado': documento,
            'doc_col_detectada':   doc_col,
            'email_col_detectada': email_col,
            'header_doc':   headers[doc_col]   if doc_col   is not None and len(headers) > doc_col   else None,
            'header_email': headers[email_col] if email_col is not None and len(headers) > email_col else None,
            'coincidencias_encontradas': len(coincidencias),
            'coincidencias': coincidencias
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/health', methods=['GET'])
def health():
    """Health check básico para verificar que el servidor está activo."""
    return jsonify({'status': 'ok'}), 200


# ==================== INICIO ====================

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    init_db()
    port  = int(os.getenv('PORT', 5001))
    debug = os.getenv('FLASK_ENV') == 'development'
    app.run(debug=debug, host='0.0.0.0', port=port)
