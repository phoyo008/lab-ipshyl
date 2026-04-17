"""
Lab Results Delivery Automation
IPS H&L Salud

Main flow:
  1. User uploads a lab result PDF (SYNLAB or COLCAN format)
  2. Patient data is extracted from the PDF (name, ID number, reference, date)
  3. Patient email is looked up in Google Sheets by ID number
  4. The PDF is sent as an email attachment via Gmail/Brevo
  5. The ENVIADO column is marked as "Si" in Google Sheets
  6. A record is saved to the local SQLite database
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
from datetime import datetime, timedelta

load_dotenv()

# ==================== CONFIGURATION ====================
# Pending sends are persisted in SQLite (table `pending_sends`) so that all
# gunicorn workers share the same state. An in-memory dict would be invisible
# across workers and cause random "Invalid or expired token" errors when
# requests are load-balanced to a different process.

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/uploads' if os.getenv('FLASK_ENV') != 'development' else 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB file size limit
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-key-change-in-production')

ALLOWED_EXTENSIONS = {'pdf'}

# Web app login credentials (configured via .env)
APP_USERNAME = os.getenv('APP_USERNAME', 'laboratorio')
APP_PASSWORD = os.getenv('APP_PASSWORD')

# Email credentials (configured via .env)
GMAIL_EMAIL    = os.getenv('GMAIL_EMAIL')
GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')

# Configurable SMTP server — defaults to Gmail, use Brevo or other on Railway
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT   = int(os.getenv('SMTP_PORT', '587'))
SMTP_USER   = os.getenv('SMTP_USER', GMAIL_EMAIL)
SMTP_PASS   = os.getenv('SMTP_PASS', GMAIL_PASSWORD)

# Google Sheets: spreadsheet name and service account credentials file
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'Directorio_IPS')
CREDENTIALS_FILE  = 'clave.json'

# SQLite database location.
#
# On Railway (and most ephemeral container platforms) /tmp is WIPED on every
# restart, redeploy, or scale event — which means the delivery history would
# disappear. To persist reports across deploys, mount a Railway Volume (or any
# persistent disk) and point DATABASE_PATH at a file on that mount, e.g.:
#
#     DATABASE_PATH=/data/reportes.db
#
# Local development still defaults to a repo-local file. The resolver below
# falls back to /tmp with a loud warning if nothing is configured, so the app
# keeps booting even when the volume is not set up yet.
def _resolve_db_path() -> str:
    explicit = os.getenv('DATABASE_PATH')
    if explicit:
        parent = os.path.dirname(explicit)
        if parent:
            os.makedirs(parent, exist_ok=True)
        return explicit
    if os.getenv('FLASK_ENV') == 'development':
        return 'reportes.db'
    print(
        "[WARNING] DATABASE_PATH is not set. Falling back to /tmp/reportes.db — "
        "delivery history WILL BE LOST on restart. "
        "Mount a Railway Volume at /data and set DATABASE_PATH=/data/reportes.db"
    )
    return '/tmp/reportes.db'


DB_PATH = _resolve_db_path()


# ==================== AUTHENTICATION ====================

def login_required(f):
    """Decorator that redirects to the login page if the user is not authenticated."""
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
        error = 'Invalid username or password'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ==================== DATABASE ====================

def init_db():
    """Create tables if they do not exist. Called at startup."""
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
    # Pending sends — shared across gunicorn workers
    conn.execute('''
        CREATE TABLE IF NOT EXISTS pending_sends (
            token       TEXT PRIMARY KEY,
            filepath    TEXT NOT NULL,
            nombre      TEXT,
            documento   TEXT,
            email       TEXT,
            referencia  TEXT,
            fecha       TEXT,
            sheet_row   INTEGER,
            created_at  TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()


# ==================== PENDING SENDS (shared across workers) ====================

def save_pending_send(token: str, data: dict) -> None:
    """Persist a pending send to SQLite."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        '''INSERT INTO pending_sends
           (token, filepath, nombre, documento, email, referencia, fecha, sheet_row, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (
            token,
            data['filepath'],
            data.get('nombre'),
            data.get('documento'),
            data.get('email'),
            data.get('referencia'),
            data.get('fecha'),
            data.get('sheet_row'),
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        ),
    )
    conn.commit()
    conn.close()


def pop_pending_send(token: str) -> dict | None:
    """Fetch a pending send and remove it atomically. Returns None if not found."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute('SELECT * FROM pending_sends WHERE token = ?', (token,)).fetchone()
    if row:
        conn.execute('DELETE FROM pending_sends WHERE token = ?', (token,))
        conn.commit()
    conn.close()
    return dict(row) if row else None


def delete_pending_send(token: str) -> dict | None:
    """Delete a pending send by token and return its row (for file cleanup)."""
    return pop_pending_send(token)


def cleanup_stale_pending_sends(max_minutes: int = 30) -> None:
    """Remove pending_sends rows older than max_minutes and delete their files."""
    cutoff = (datetime.now() - timedelta(minutes=max_minutes)).strftime('%Y-%m-%d %H:%M:%S')
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    stale = conn.execute(
        'SELECT token, filepath FROM pending_sends WHERE created_at < ?',
        (cutoff,),
    ).fetchall()
    conn.execute('DELETE FROM pending_sends WHERE created_at < ?', (cutoff,))
    conn.commit()
    conn.close()
    for r in stale:
        fp = r['filepath']
        if fp and os.path.exists(fp):
            try:
                os.remove(fp)
            except OSError:
                pass


def save_to_db(nombre, documento, email, referencia, fecha_resultado, estado):
    """
    Insert a record into the 'envios' table.
    The estado (status) field can be 'Enviado' (sent) or an error message like 'Error: ...'.
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
        print(f"Error saving to DB: {e}")


# ==================== HELPER FUNCTIONS ====================

def allowed_file(filename):
    """Check that the file has a .pdf extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_sheets_client():
    """
    Return an authenticated gspread client using the service account.
    Reads credentials from the GOOGLE_CREDENTIALS_JSON environment variable
    (recommended for cloud production) or from the clave.json file (local development).
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


# ==================== PDF DATA EXTRACTION ====================

def extract_pdf_data(filepath):
    """
    Extract patient data from a lab result PDF.

    Supported lab formats:
      - SYNLAB:  NOMBRE / DOCUMENTO: CC. / REFERENCIA / FECHA INGRESO
      - COLCAN:  Nombre / Idenficacion: CC / numeric barcode / Fecha toma muestra

    Note on COLCAN: their defective font encodes the letter 't' as a null byte
    (\\x00). The text is cleaned before applying regex patterns.

    Returns a dict with keys: nombre, documento, referencia, fecha, resultado_completo.
    Returns None if the PDF cannot be opened.
    """
    try:
        with pdfplumber.open(filepath) as pdf:
            text = ''.join((page.extract_text() or '') + '\n' for page in pdf.pages)

        # Fix defective encoding in COLCAN PDFs:
        # pdfplumber extracts \x00 where 't' should be
        text = text.replace('\x00', 't')

        # NAME — stops before Nro, REFERENCIA, DOCUMENTO, Iden, or Tel:
        nombre_match = re.search(
            r'(?:NOMBRE|Nombre)\s*:\s*(.+?)(?=\s*(?:N[°º]|REFERENCIA|DOCUMENTO|Iden|Tel:|$))',
            text, re.IGNORECASE
        )

        # ID NUMBER — "Iden\w+" covers both "Identificacion" and "Idenficacion" (COLCAN typo)
        documento_match = re.search(
            r'(?:DOCUMENTO|Iden\w+)\s*:\s*CC\.?\s*(\d{5,12})',
            text, re.IGNORECASE
        )

        # REFERENCE — SYNLAB uses "REFERENCIA: 123"; COLCAN has a numeric barcode on its own line
        referencia_match = re.search(r'REFERENCIA\s*:\s*(\d+)', text, re.IGNORECASE)
        if not referencia_match:
            # Fallback: look for a 10-15 digit number on its own line (COLCAN barcode)
            referencia_match = re.search(r'^\s*(\d{10,15})\s*$', text, re.MULTILINE)

        # DATE — accepts both SYNLAB and COLCAN date formats
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
        print(f"Error extracting PDF: {e}")
        return None


# ==================== GOOGLE SHEETS ====================

def _parse_sheet_date(value):
    """
    Convert a date/timestamp value from the sheet to a datetime object for comparison.
    Supports common Google Forms formats:
      - 'DD/MM/YYYY HH:MM:SS'  (Forms timestamp)
      - 'DD/MM/YYYY'
      - 'YYYY-MM-DD HH:MM:SS'
      - 'YYYY-MM-DD'
    Returns datetime or datetime.min if parsing fails (so it sorts to the bottom).
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
    Look up a patient's email in Google Sheets by their ID number.

    When multiple rows match the same ID (patient with multiple visits),
    the row with the most recent date is selected using the timestamp column
    ('Marca temporal' from Google Forms or another detected date column).

    Sheet structure:
      - Row 1: Internal IDs (ignored)
      - Row 2: Column headers
      - Row 3+: Patient data

    Columns detected by header name:
      - Date/timestamp: column containing 'marca' and 'temporal', or 'fecha'
      - ID number: column containing 'numero' and 'documento'
      - Results email: column containing 'e-mail' and 'resultado'

    Returns: (email, row_number) or (None, None) if not found.
    """
    try:
        client = get_sheets_client()
        sheet  = client.open(GOOGLE_SHEET_NAME).sheet1
        rows   = sheet.get_all_values()

        # Row 2 (index 1) contains the headers
        headers = rows[1]

        doc_col   = next((i for i, h in enumerate(headers) if 'número' in h.lower() and 'documento' in h.lower()), None)
        email_col = next((i for i, h in enumerate(headers) if 'e-mail' in h.lower() and 'resultado' in h.lower()), None)

        # Date column: prefer 'Marca temporal' (Google Forms), otherwise any column with 'fecha'
        fecha_col = next((i for i, h in enumerate(headers) if 'marca' in h.lower() and 'temporal' in h.lower()), None)
        if fecha_col is None:
            fecha_col = next((i for i, h in enumerate(headers) if 'fecha' in h.lower()), None)

        if doc_col is None or email_col is None:
            print("Required columns not found:", headers)
            return None, None

        # Collect ALL rows matching the ID number
        matches = []
        for idx, row in enumerate(rows[2:]):
            if len(row) > doc_col and str(row[doc_col]).strip() == str(documento).strip():
                sheet_row = idx + 3  # Actual row number (1-based) in the sheet
                email     = row[email_col].strip() if len(row) > email_col else ''

                fecha_val = ''
                if fecha_col is not None and len(row) > fecha_col:
                    fecha_val = row[fecha_col]

                matches.append({
                    'email':      email,
                    'sheet_row':  sheet_row,
                    'fecha_dt':   _parse_sheet_date(fecha_val),
                    'fecha_raw':  fecha_val
                })

        if not matches:
            return None, None

        if len(matches) > 1:
            print(f"Document {documento}: {len(matches)} records found.")
            for m in matches:
                print(f"  -> row {m['sheet_row']} | date: {m['fecha_raw']} | email: '{m['email']}'")

        # Prefer rows that have an email; among those, pick the most recent
        with_email = [m for m in matches if m['email']]

        if with_email:
            best = max(with_email, key=lambda x: x['fecha_dt'])
            print(f"  -> Selected row {best['sheet_row']} (date: {best['fecha_raw']}, email: {best['email']})")
            return best['email'], best['sheet_row']
        else:
            # Patient found but no row has an email
            # Return the most recent row so we can mark "No email" in the sheet
            no_email = max(matches, key=lambda x: x['fecha_dt'])
            print(f"Document {documento}: found in row {no_email['sheet_row']} but no email on file.")
            return None, no_email['sheet_row']

    except Exception as e:
        print(f"Error searching Google Sheets: {e}")
        return None, None


def mark_sent_in_sheet(sheet_row):
    """Update the 'ENVIADO' column to 'Si' in the specified Google Sheet row."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        headers = sheet.row_values(2)  # Headers in row 2

        # Find the ENVIADO column index (1-based for update_cell)
        enviado_col = next((i + 1 for i, h in enumerate(headers) if h.strip().upper() == 'ENVIADO'), None)

        if enviado_col is None:
            print("ENVIADO column not found")
            return

        sheet.update_cell(sheet_row, enviado_col, 'Si')

    except Exception as e:
        print(f"Error marking as sent in sheet: {e}")


def mark_no_email_in_sheet(sheet_row):
    """Mark 'Sin correo' (no email) in the ENVIADO column when the patient has no registered email."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        headers = sheet.row_values(2)

        enviado_col = next((i + 1 for i, h in enumerate(headers) if h.strip().upper() == 'ENVIADO'), None)

        if enviado_col is None:
            print("ENVIADO column not found")
            return

        sheet.update_cell(sheet_row, enviado_col, 'Sin correo')

    except Exception as e:
        print(f"Error marking no-email in sheet: {e}")


# ==================== EMAIL SENDING ====================

def _build_email_body(patient_name: str, items: list) -> tuple[str, str]:
    """
    Build the subject and plain-text body for a lab result email.
    `items` is a list of dicts with keys: referencia, fecha, pdf_path.
    Returns (subject, body).
    """
    name = patient_name.strip() if patient_name else ''
    if len(items) == 1:
        it = items[0]
        subject = f"Lab Results - Ref: {it['referencia']}"
        body = (
            f"Dear {name},\n\n"
            f"Please find your lab results attached.\n\n"
            f"REFERENCE: {it['referencia']}\n"
            f"DATE:      {it['fecha']}\n\n"
            f"If you have any questions about your results, please contact your physician.\n\n"
            f"Best regards,\n"
            f"IPS H&L Salud - Laboratory"
        )
        return subject, body

    lines = "\n".join(
        f"  - Reference: {it['referencia']}  |  Date: {it['fecha']}" for it in items
    )
    subject = f"Lab Results - {len(items)} results attached"
    body = (
        f"Dear {name},\n\n"
        f"Please find your {len(items)} lab results attached:\n\n"
        f"{lines}\n\n"
        f"If you have any questions about your results, please contact your physician.\n\n"
        f"Best regards,\n"
        f"IPS H&L Salud - Laboratory"
    )
    return subject, body


def send_email(to_email: str, patient_name: str, items: list) -> bool:
    """
    Send a single email with one or more lab result PDFs attached.

    `items` is a list of dicts with keys: referencia, fecha, pdf_path.
    All items in a call are delivered together as attachments on the same
    message — this lets us group multiple results for one patient into a
    single email instead of spamming them with one per PDF.

    Uses Brevo API if BREVO_API_KEY is configured (production on Railway — HTTPS).
    Falls back to Gmail SMTP if not configured (local development with App Password).
    Returns True if the email was sent successfully, False otherwise.
    """
    if not items:
        return False

    subject, body = _build_email_body(patient_name, items)
    brevo_key = os.getenv('BREVO_API_KEY')

    if brevo_key:
        # ---- Brevo API (production on Railway — uses HTTPS, not SMTP) ----
        try:
            import base64, requests as req
            attachments = []
            for it in items:
                with open(it['pdf_path'], 'rb') as f:
                    pdf_data = base64.b64encode(f.read()).decode()
                attachments.append({
                    "content": pdf_data,
                    "name": f"result_{it['referencia']}.pdf",
                })

            payload = {
                "sender":      {"name": "IPS H&L Salud - Laboratory", "email": GMAIL_EMAIL},
                "to":          [{"email": to_email}],
                "subject":     subject,
                "textContent": body,
                "attachment":  attachments,
            }
            response = req.post(
                "https://api.brevo.com/v3/smtp/email",
                headers={"api-key": brevo_key, "Content-Type": "application/json"},
                json=payload,
                timeout=30,
            )
            if response.status_code == 201:
                return True
            print(f"Brevo API error: {response.status_code} {response.text}")
            return False

        except Exception as e:
            print(f"Error sending email (Brevo API): {e}")
            return False

    # ---- SMTP (local Gmail with App Password) ----
    try:
        message = MIMEMultipart()
        message['From']    = GMAIL_EMAIL
        message['To']      = to_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        for it in items:
            with open(it['pdf_path'], 'rb') as f:
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition',
                f'attachment; filename="result_{it["referencia"]}.pdf"',
            )
            message.attach(attachment)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(message)
        server.quit()
        return True

    except Exception as e:
        print(f"Error sending email (SMTP): {e}")
        return False


def log_error(patient_name, documento, error_msg):
    """Log an error to the local database."""
    save_to_db(patient_name, documento, '', 'N/A', 'N/A', f'Error: {error_msg}')


# ==================== ROUTES ====================

@app.route('/')
@login_required
def index():
    """Main page: PDF upload form."""
    return render_template('index.html')


@app.route('/reportes')
@login_required
def reportes():
    """
    History page: displays all delivery records from SQLite,
    ordered from most recent to oldest.
    """
    conn   = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    envios = conn.execute('SELECT * FROM envios ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('reportes.html', envios=envios)


@app.route('/api/preview-pdf', methods=['POST'])
@login_required
def preview_pdf():
    """
    STEP 1 of the delivery flow.
    Receives the PDF, extracts patient data, and looks up their email in Google Sheets.
    Does NOT send any email — only returns information for the user to confirm.

    Returns a token identifying the pending delivery, along with patient data.
    The token must be sent to /api/confirm-send to execute the actual delivery.
    """
    filepath = None
    try:
        cleanup_stale_pending_sends()

        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        if not allowed_file(file.filename):
            return jsonify({'error': 'Only PDF files are allowed'}), 400

        # Save the PDF — kept until the user confirms or cancels
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        filename = f"{uuid.uuid4().hex}_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        patient_data = extract_pdf_data(filepath)
        if not patient_data or not patient_data.get('documento'):
            log_error('Unknown', 'N/A', 'Could not extract data from PDF')
            return jsonify({'error': 'Could not extract data from PDF'}), 400

        nombre     = patient_data.get('nombre', 'Unknown')
        documento  = patient_data.get('documento', 'N/A')
        referencia = patient_data.get('referencia', 'N/A')
        fecha      = patient_data.get('fecha', 'N/A')

        # Look up email in Google Sheets (picks the row with the most recent date)
        patient_email, sheet_row = find_email_in_sheets(documento)
        if not patient_email:
            if sheet_row:
                mark_no_email_in_sheet(sheet_row)
                save_to_db(nombre, documento, '', referencia, fecha, 'Sin correo')
                return jsonify({
                    'error': 'Patient has no registered email. Marked as "No email" in the directory.',
                    'nombre': nombre,
                    'documento': documento
                }), 404
            else:
                log_error(nombre, documento, 'Patient not found in directory')
                return jsonify({
                    'error': 'Patient not found in directory',
                    'nombre': nombre,
                    'documento': documento
                }), 404

        # Validate email format
        if not re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', patient_email):
            log_error(nombre, documento, f'Invalid email: {patient_email}')
            return jsonify({'error': f'Invalid email: {patient_email}', 'nombre': nombre}), 400

        # Generate token and persist pending state (shared across workers)
        token = uuid.uuid4().hex
        save_pending_send(token, {
            'filepath':   filepath,
            'nombre':     nombre,
            'documento':  documento,
            'email':      patient_email,
            'referencia': referencia,
            'fecha':      fecha,
            'sheet_row':  sheet_row,
        })

        return jsonify({
            'token':     token,
            'nombre':    nombre,
            'documento': documento,
            'email':     patient_email,
            'referencia': referencia,
            'fecha':     fecha
        }), 200

    except Exception as e:
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': f'Error: {str(e)}'}), 500


@app.route('/api/confirm-send', methods=['POST'])
@login_required
def confirm_send():
    """
    STEP 2 of the delivery flow.

    Accepts either:
      - {"token": "<token>"}       (single delivery, kept for backward compat)
      - {"tokens": ["<t1>", ...]}  (batch — preferred; groups results by patient
                                    so multiple PDFs for the same documento are
                                    attached to a single email)

    For each group (keyed by documento):
      1. Sends ONE email with all PDFs attached
      2. Marks ENVIADO=Si in Google Sheets (once per patient)
      3. Saves a record in SQLite per PDF
      4. Deletes the temporary PDFs
    """
    try:
        data = request.get_json() or {}

        tokens = data.get('tokens')
        if not tokens:
            single = data.get('token')
            tokens = [single] if single else []

        if not tokens:
            return jsonify({'error': 'No tokens provided'}), 400

        # Fetch and remove each pending send. Missing tokens are reported back.
        entries  = []
        missing  = []
        for t in tokens:
            entry = pop_pending_send(t)
            if entry:
                entries.append(entry)
            else:
                missing.append(t)

        if not entries:
            return jsonify({
                'error': 'Invalid or expired tokens. Please re-upload the PDFs.',
                'missing_tokens': missing
            }), 400

        # Group by documento so each patient receives a single email.
        groups: dict = {}
        for e in entries:
            groups.setdefault(e['documento'], []).append(e)

        sent_groups = []
        errors      = []
        files_to_cleanup = [e['filepath'] for e in entries]

        for documento, group in groups.items():
            email     = group[0]['email']
            nombre    = group[0]['nombre']
            sheet_row = group[0]['sheet_row']

            items = [
                {
                    'referencia': g['referencia'],
                    'fecha':      g['fecha'],
                    'pdf_path':   g['filepath'],
                }
                for g in group
            ]

            ok = send_email(to_email=email, patient_name=nombre, items=items)

            if ok:
                if sheet_row:
                    mark_sent_in_sheet(sheet_row)
                for g in group:
                    save_to_db(g['nombre'], g['documento'], email,
                               g['referencia'], g['fecha'], 'Enviado')
                sent_groups.append({
                    'nombre':      nombre,
                    'documento':   documento,
                    'email':       email,
                    'count':       len(group),
                    'referencias': [g['referencia'] for g in group],
                })
            else:
                for g in group:
                    log_error(g['nombre'], g['documento'], 'Failed to send email')
                errors.append({
                    'nombre':    nombre,
                    'documento': documento,
                    'email':     email,
                    'count':     len(group),
                    'error':     'Failed to send email. Check email credentials.',
                })

        # Clean up temp PDFs regardless of outcome
        for fp in files_to_cleanup:
            if fp and os.path.exists(fp):
                try:
                    os.remove(fp)
                except OSError:
                    pass

        status_code = 200 if sent_groups else 500
        return jsonify({
            'success':      len(errors) == 0 and bool(sent_groups),
            'sent_groups':  sent_groups,
            'errors':       errors,
            'missing_tokens': missing,
        }), status_code

    except Exception as e:
        return jsonify({'error': f'Error: {str(e)}'}), 500


@app.route('/api/cancel-send', methods=['POST'])
@login_required
def cancel_send():
    """Cancel a pending delivery and delete the temporary PDF."""
    try:
        data  = request.get_json()
        token = data.get('token') if data else None

        if token:
            entry = pop_pending_send(token)
            if entry:
                fp = entry.get('filepath')
                if fp and os.path.exists(fp):
                    try:
                        os.remove(fp)
                    except OSError:
                        pass

        return jsonify({'ok': True}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/test-connection', methods=['GET'])
@login_required
def test_connection():
    """
    Test external connections:
      - Google Sheets: attempts to open the sheet and count records
      - Email (Brevo API or Gmail SMTP): attempts to authenticate

    Returns 200 if everything is OK, 500 if any connection fails.
    """
    results = {}

    # Test Google Sheets
    try:
        client = get_sheets_client()
        sheet  = client.open(GOOGLE_SHEET_NAME).sheet1
        rows   = sheet.get_all_values()
        total  = len(rows) - 2 if len(rows) > 2 else 0  # Subtract ID row and headers
        results['google_sheets'] = {
            'ok':      True,
            'mensaje': f'Connected. {total} records in directory.',
            'hoja':    GOOGLE_SHEET_NAME
        }
    except Exception as e:
        results['google_sheets'] = {'ok': False, 'mensaje': str(e)}

    # Test email: Brevo API (production) or Gmail SMTP (local)
    brevo_key = os.getenv('BREVO_API_KEY')
    try:
        if brevo_key:
            import requests as req
            r = req.get('https://api.brevo.com/v3/account',
                        headers={'api-key': brevo_key}, timeout=10)
            if r.status_code == 200:
                results['email'] = {'ok': True, 'mensaje': f'Brevo API connected. Sending from {GMAIL_EMAIL}'}
            else:
                raise ValueError(f'Invalid Brevo API key: {r.status_code}')
        else:
            if not SMTP_USER:
                raise ValueError('Email credentials not configured')
            servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10)
            servidor.starttls()
            servidor.login(SMTP_USER, SMTP_PASS)
            servidor.quit()
            results['email'] = {'ok': True, 'mensaje': f'Gmail SMTP connected as {SMTP_USER}'}
    except Exception as e:
        results['email'] = {'ok': False, 'mensaje': str(e)}

    all_ok = all(r['ok'] for r in results.values())
    return jsonify({'ok': all_ok, 'resultados': results}), 200 if all_ok else 500


@app.route('/api/debug-sheet', methods=['GET'])
@login_required
def debug_sheet():
    """Show actual Google Sheet headers for diagnosing column detection issues."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        rows    = sheet.get_all_values()
        headers = rows[1] if len(rows) > 1 else []
        return jsonify({
            'total_rows': len(rows),
            'row2_headers': {str(i): h for i, h in enumerate(headers)}
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/debug-search/<documento>', methods=['GET'])
@login_required
def debug_search(documento):
    """Search for a document ID directly in the sheet and show detected columns and matching rows."""
    try:
        client  = get_sheets_client()
        sheet   = client.open(GOOGLE_SHEET_NAME).sheet1
        rows    = sheet.get_all_values()
        headers = rows[1] if len(rows) > 1 else []

        doc_col   = next((i for i, h in enumerate(headers) if 'número' in h.lower() and 'documento' in h.lower()), None)
        email_col = next((i for i, h in enumerate(headers) if 'e-mail' in h.lower() and 'resultado' in h.lower()), None)

        matches = []
        for idx, row in enumerate(rows[2:]):
            if len(row) > (doc_col or 0):
                val = str(row[doc_col]).strip() if doc_col is not None else ''
                if val == str(documento).strip():
                    matches.append({
                        'sheet_row': idx + 3,
                        'doc_value':  val,
                        'email':      row[email_col].strip() if email_col is not None and len(row) > email_col else 'COL_OUT_OF_RANGE',
                        'nombre':     row[6].strip() if len(row) > 6 else ''
                    })

        return jsonify({
            'searched_document': documento,
            'detected_doc_col':   doc_col,
            'detected_email_col': email_col,
            'header_doc':   headers[doc_col]   if doc_col   is not None and len(headers) > doc_col   else None,
            'header_email': headers[email_col] if email_col is not None and len(headers) > email_col else None,
            'matches_found': len(matches),
            'matches': matches
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/health', methods=['GET'])
def health():
    """Basic health check to verify the server is running."""
    return jsonify({'status': 'ok'}), 200


# ==================== STARTUP ====================

# Initialize DB always, even when the server starts via gunicorn
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
init_db()

if __name__ == '__main__':
    port  = int(os.getenv('PORT', 5001))
    debug = os.getenv('FLASK_ENV') == 'development'
    app.run(debug=debug, host='0.0.0.0', port=port)
