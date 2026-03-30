# Automatizador de Resultados de Laboratorio — IPS H&L Salud

Aplicación web en Flask que automatiza el envío de resultados de laboratorio por correo electrónico a los pacientes.

## Cómo funciona

1. El operario sube el PDF del resultado desde la interfaz web
2. La app extrae los datos del paciente del PDF (nombre, documento, referencia, fecha)
3. Busca el email del paciente en Google Sheets por número de documento
4. Envía el PDF como adjunto al correo del paciente vía Gmail
5. Marca la fila del paciente como `ENVIADO = Si` en Google Sheets
6. Guarda un registro del envío en la base de datos local SQLite

Laboratorios soportados: **SYNLAB** y **COLCAN**

---

## Estructura del proyecto

```
automatizacion_ips/
├── app.py                  # Aplicación principal Flask
├── clave.json              # Credenciales del service account de Google (NO compartir)
├── .env                    # Variables de entorno (NO subir a git)
├── requirements.txt        # Dependencias de Python
├── reportes.db             # Base de datos SQLite (se genera automáticamente)
├── uploads/                # Carpeta temporal para PDFs subidos (se limpia automáticamente)
├── templates/
│   ├── index.html          # Página principal (formulario de carga)
│   └── reportes.html       # Historial de envíos
└── static/
    ├── style.css           # Estilos globales
    └── logo.png            # Logo de IPS H&L Salud
```

---

## Requisitos previos

- Python 3.9 o superior
- Cuenta de Gmail con **App Password** habilitada (autenticación en dos pasos requerida)
- Cuenta de Google Cloud con un **Service Account** y las APIs de Google Sheets y Google Drive activadas
- El service account debe tener acceso de editor a la hoja de Google Sheets

---

## Instalación

```bash
# 1. Crear y activar entorno virtual
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Configurar variables de entorno (ver sección siguiente)
cp .env.example .env
# Editar .env con tus credenciales

# 4. Colocar clave.json en la raíz del proyecto

# 5. Iniciar el servidor
python app.py
```

La app quedará disponible en: `http://localhost:5001`

---

## Configuración (.env)

```env
# Gmail — usar App Password, no la contraseña normal de la cuenta
GMAIL_EMAIL=tu_correo@gmail.com
GMAIL_PASSWORD=xxxx xxxx xxxx xxxx

# Google Sheets — nombre exacto de la hoja (como aparece en Google Drive)
GOOGLE_SHEET_NAME=NOMBRE_DE_TU_HOJA

# Flask
FLASK_ENV=development
SECRET_KEY=cambia_esto_por_una_clave_aleatoria
PORT=5001
```

### Cómo obtener el App Password de Gmail

1. Ir a [myaccount.google.com](https://myaccount.google.com)
2. Seguridad → Verificación en dos pasos (activar si no está activa)
3. Seguridad → Contraseñas de aplicaciones
4. Crear una nueva para "Correo" / "Otro (nombre personalizado)"
5. Copiar la clave de 16 caracteres al `.env`

---

## Configuración de Google Sheets

### Estructura esperada de la hoja

| (fila 1) | IDs internos | ... |
|---|---|---|
| (fila 2) | **Marca temporal** | **Número de documento** | **E-mail resultado** | **ENVIADO** | ... |
| (fila 3+) | 26/02/2026 10:30:00 | 12345678 | paciente@email.com | | ... |

- La fila 1 puede contener cualquier cosa (IDs de formulario, etc.)
- La fila 2 debe contener los encabezados de columna
- La columna de fecha preferida es `Marca temporal` (Google Forms la agrega automáticamente). Si no existe, se usa cualquier columna que contenga "fecha" en el nombre
- La columna de documento debe tener "número" y "documento" en su nombre
- La columna de email debe tener "e-mail" y "resultado" en su nombre
- La columna ENVIADO debe llamarse exactamente `ENVIADO`

**Pacientes con múltiples registros:** cuando el mismo número de documento aparece en varias filas (paciente con varias visitas), la app selecciona automáticamente la fila con la fecha más reciente y usa el email de esa fila.

### Cómo configurar el Service Account

1. Ir a [Google Cloud Console](https://console.cloud.google.com)
2. Crear un proyecto o usar uno existente
3. Activar las APIs: **Google Sheets API** y **Google Drive API**
4. Crear credenciales → Cuenta de servicio
5. Descargar el archivo JSON y guardarlo como `clave.json` en la raíz del proyecto
6. En Google Sheets, compartir la hoja con el email del service account (`...@...iam.gserviceaccount.com`) con permisos de **Editor**

---

## Rutas de la API

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/` | Página principal con formulario |
| GET | `/reportes` | Historial de envíos |
| POST | `/api/process-pdf` | Procesa y envía un PDF |
| GET | `/api/test-connection` | Verifica conexión con Sheets y Gmail |
| GET | `/api/health` | Health check del servidor |

---

## Detalles técnicos de extracción PDF

### SYNLAB
- Nombre: campo `NOMBRE:`
- Documento: campo `DOCUMENTO: CC.`
- Referencia: campo `REFERENCIA:`
- Fecha: campo `FECHA INGRESO:`

### COLCAN
- Nombre: campo `Nombre:`
- Documento: campo `Idenficación: CC` *(typo en el PDF original)*
- Referencia: código de barras numérico de 10-15 dígitos en su propia línea
- Fecha: campo `Fecha toma muestra:`
- **Nota:** los PDFs de COLCAN usan una fuente defectuosa que codifica la letra `t` como byte nulo (`\x00`). La app corrige esto automáticamente antes de procesar el texto.

---

## Dependencias principales

| Paquete | Uso |
|---------|-----|
| Flask | Framework web |
| pdfplumber | Extracción de texto de PDFs |
| gspread | Cliente de Google Sheets |
| google-auth | Autenticación con service account |
| python-dotenv | Carga de variables de entorno |
| gunicorn | Servidor WSGI para producción |

---

## Producción

Para correr en producción usar gunicorn en lugar del servidor de desarrollo de Flask:

```bash
gunicorn -w 2 -b 0.0.0.0:5001 app:app
```

Cambiar también en `.env`:
```env
FLASK_ENV=production
SECRET_KEY=clave_aleatoria_larga_y_segura
```
