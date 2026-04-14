# Lab Results Delivery Automation

A Flask web application that automates sending laboratory test results to patients via email. Built for **IPS H&L Salud**, a healthcare clinic in Colombia.

## What It Does

Staff upload a lab result PDF through a simple web interface. The system automatically:

1. **Extracts patient data** from the PDF using regex-based parsing (name, ID number, reference, date)
2. **Looks up the patient's email** in a Google Sheets directory by document number
3. **Previews the delivery** for staff confirmation before sending
4. **Sends the PDF** as an email attachment via Gmail SMTP or Brevo API
5. **Marks the delivery** as complete in Google Sheets
6. **Logs everything** to a local SQLite database for audit

Supports batch processing — multiple PDFs can be uploaded and sent at once.

## Supported Lab Formats

| Lab | Name Field | ID Field | Reference | Date Field |
|-----|-----------|----------|-----------|------------|
| **SYNLAB** | `NOMBRE:` | `DOCUMENTO: CC.` | `REFERENCIA:` | `FECHA INGRESO:` |
| **COLCAN** | `Nombre:` | `Idenficacion: CC` | Numeric barcode (10-15 digits) | `Fecha toma muestra:` |

> COLCAN PDFs use a defective font that encodes the letter `t` as a null byte (`\x00`). The parser corrects this automatically.

## Architecture

```
┌─────────────┐     ┌──────────────┐     ┌──────────────┐
│  Browser UI  │────▶│  Flask API   │────▶│ Google Sheets│
│  (upload PDF)│◀────│  (Python)    │◀────│ (patient dir)│
└─────────────┘     │              │     └──────────────┘
                    │  ┌────────┐  │     ┌──────────────┐
                    │  │pdfplumber│ │────▶│  Gmail/Brevo │
                    │  │(extract)│  │     │  (send email)│
                    │  └────────┘  │     └──────────────┘
                    │  ┌────────┐  │
                    │  │ SQLite  │  │
                    │  │ (audit) │  │
                    │  └────────┘  │
                    └──────────────┘
```

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Backend | Python 3.12, Flask |
| PDF parsing | pdfplumber + regex |
| Patient directory | Google Sheets via gspread |
| Email delivery | Gmail SMTP or Brevo API |
| Database | SQLite |
| Auth | Session-based login |
| Deployment | Gunicorn, Railway-ready |

## Quick Start

### Prerequisites

- Python 3.9+
- A Gmail account with [App Password](https://myaccount.google.com/apppasswords) enabled (requires 2-Step Verification)
- A Google Cloud [Service Account](https://console.cloud.google.com/iam-admin/serviceaccounts) with Sheets and Drive APIs enabled

### Setup

```bash
# Clone the repository
git clone https://github.com/phoyo008/laboratorio-ipshyl.git
cd laboratorio-ipshyl

# Create and activate a virtual environment
python -m venv venv
source venv/bin/activate  # macOS/Linux
# venv\Scripts\activate   # Windows

# Install dependencies
pip install -r requirements.txt

# Configure environment variables
cp .env.example .env
# Edit .env with your credentials

# Place your Google service account key as clave.json in the project root

# Start the server
python app.py
```

The app will be available at `http://localhost:5001`.

### Google Sheets Setup

The spreadsheet should have this structure:

| Row 1 | *(internal IDs — ignored)* |
|-------|---------------------------|
| **Row 2** | **Marca temporal** | **Numero de documento** | **E-mail resultado** | **ENVIADO** |
| Row 3+ | `26/02/2026 10:30:00` | `12345678` | `patient@email.com` | |

- **Row 2** must contain the column headers
- The document column must include "numero" and "documento" in its name
- The email column must include "e-mail" and "resultado" in its name
- The `ENVIADO` column must be named exactly `ENVIADO`
- When a patient has multiple rows (repeat visits), the most recent entry is used

Share the spreadsheet with the service account email (`...@...iam.gserviceaccount.com`) with **Editor** permissions.

## API Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| `GET` | `/` | Main page — PDF upload form |
| `GET` | `/reportes` | Delivery history and stats |
| `POST` | `/api/preview-pdf` | Extract data from PDF, look up patient email |
| `POST` | `/api/confirm-send` | Send the email (requires preview token) |
| `POST` | `/api/cancel-send` | Cancel a pending delivery |
| `GET` | `/api/test-connection` | Test Google Sheets and email connectivity |
| `GET` | `/api/health` | Health check |
| `GET` | `/api/debug-sheet` | Show detected sheet headers |
| `GET` | `/api/debug-search/<id>` | Debug patient lookup by document ID |

## Production Deployment

The app is Railway-ready with the included `Procfile` and `runtime.txt`.

```bash
# Run with gunicorn
gunicorn -w 2 -b 0.0.0.0:$PORT app:app
```

For production, set these environment variables:
- `FLASK_ENV=production`
- `SECRET_KEY` — a long random string
- `GOOGLE_CREDENTIALS_JSON` — service account JSON (instead of the `clave.json` file)
- `BREVO_API_KEY` — for email delivery via Brevo instead of Gmail SMTP

## License

MIT
