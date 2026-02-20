from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from flask_wtf.csrf import CSRFProtect, CSRFError
from dotenv import load_dotenv
import os, uuid, time
from pathlib import Path

# Load environment variables from .env file and prefer local project values.
# This avoids machine-level env vars unexpectedly overriding local CSRF/session settings.
load_dotenv(override=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.getenv("UPLOAD_FOLDER", "uploads")
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_CONTENT_LENGTH", 32 * 1024 * 1024))
app.config["MAX_FILE_SIZE"] = int(os.getenv("MAX_FILE_SIZE", 32 * 1024 * 1024))
app.config["FILE_EXPIRY_HOURS"] = int(os.getenv("FILE_EXPIRY_HOURS", 24))
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-key-change-in-production")
session_cookie_secure = os.getenv("SESSION_COOKIE_SECURE")
if session_cookie_secure is None:
    # Default secure cookies in production, allow HTTP during local development.
    app.config["SESSION_COOKIE_SECURE"] = os.getenv("FLASK_ENV", "production").lower() == "production"
else:
    app.config["SESSION_COOKIE_SECURE"] = session_cookie_secure.lower() == "true"
app.config["SESSION_COOKIE_HTTPONLY"] = os.getenv("SESSION_COOKIE_HTTPONLY", "True").lower() == "true"
app.config["SESSION_COOKIE_SAMESITE"] = os.getenv("SESSION_COOKIE_SAMESITE", "Lax")

# Enable CSRF protection
csrf = CSRFProtect(app)

# CSRF error handler
@app.errorhandler(CSRFError)
def handle_csrf_error(e):
    payload = {"error": "CSRF token validation failed. Please refresh and try again."}
    if app.debug or os.getenv("CSRF_ERROR_DETAILS", "False").lower() == "true":
        payload["details"] = e.description
    return jsonify(payload), 400

os.makedirs("uploads", exist_ok=True)

_last_cleanup = 0.0
_CLEANUP_INTERVAL = 300  # run cleanup at most every 5 minutes

def cleanup_old_files(max_age_hours=None):
    """
    Remove files older than max_age_hours.
    Called periodically to free up disk space.
    """
    if max_age_hours is None:
        max_age_hours = app.config["FILE_EXPIRY_HOURS"]
    
    upload_folder = Path(app.config["UPLOAD_FOLDER"])
    if not upload_folder.exists():
        return
    
    current_time = time.time()
    max_age_seconds = max_age_hours * 3600
    
    try:
        for file_path in upload_folder.glob("*"):
            if file_path.is_file():
                file_age = current_time - file_path.stat().st_mtime
                if file_age > max_age_seconds:
                    file_path.unlink()
                    print(f"Cleaned up old file: {file_path.name}")
    except Exception as e:
        print(f"Error during file cleanup: {e}")

@app.before_request
def cleanup_before_request():
    """Run cleanup at most every _CLEANUP_INTERVAL seconds to avoid blocking every request."""
    global _last_cleanup
    now = time.time()
    if now - _last_cleanup >= _CLEANUP_INTERVAL:
        _last_cleanup = now
        cleanup_old_files()

def pdf_to_word(src, out):
    from pdf2docx import Converter
    cv = Converter(src)
    cv.convert(out)
    cv.close()

def pdf_to_excel(src, out):
    import pdfplumber, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    with pdfplumber.open(src) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        ws.append([c if c else "" for c in row])
            else:
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        ws.append([line])
    wb.save(out)

def word_to_pdf(src, out):
    from docx2pdf import convert
    convert(src, out)

def excel_to_pdf(src, out):
    import openpyxl
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    
    # Load workbook
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    
    # Create PDF
    doc = SimpleDocTemplate(out, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    
    # Extract data from worksheet
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append([str(cell) if cell is not None else "" for cell in row])
    
    if data:
        # Fit all columns across the page width
        col_count = len(data[0])
        page_width = A4[0] - 1.0 * inch  # usable width after margins
        col_width = min(1.5 * inch, page_width / col_count) if col_count else 1.5 * inch
        table = Table(data, colWidths=[col_width] * col_count)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)
    
    doc.build(elements)

MODES = {
    "pdf-to-word":  {"fn": pdf_to_word,  "ext": ".docx", "input_ext": [".pdf"]},
    "pdf-to-excel": {"fn": pdf_to_excel, "ext": ".xlsx", "input_ext": [".pdf"]},
    "word-to-pdf":  {"fn": word_to_pdf,  "ext": ".pdf", "input_ext": [".docx", ".doc"]},
    "excel-to-pdf": {"fn": excel_to_pdf, "ext": ".pdf", "input_ext": [".xlsx", ".xls"]},
}

# MIME types to allow for each file type
ALLOWED_MIME_TYPES = {
    ".pdf": ["application/pdf"],
    ".docx": ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    ".doc": ["application/msword"],
    ".xlsx": ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    ".xls": ["application/vnd.ms-excel"],
}

def validate_file(file, allowed_extensions, max_size):
    """
    Validate file type, extension, and size.
    Returns (is_valid, error_message)
    """
    if not file or file.filename == "":
        return False, "No file provided."
    
    # Get file extension
    file_ext = os.path.splitext(file.filename)[1].lower()
    
    # Check extension
    if file_ext not in allowed_extensions:
        return False, f"Invalid file type. Allowed types: {', '.join(allowed_extensions)}"
    
    # Check MIME type
    if hasattr(file, 'content_type'):
        allowed_mimes = ALLOWED_MIME_TYPES.get(file_ext, [])
        if allowed_mimes and file.content_type not in allowed_mimes:
            return False, f"Invalid file format for {file_ext} file."
    
    # Check file size (seek to end and get position)
    file.seek(0, 2)  # Seek to end
    file_size = file.tell()
    file.seek(0)  # Reset to beginning
    
    if file_size > max_size:
        max_mb = max_size / (1024 * 1024)
        return False, f"File too large. Maximum size is {max_mb:.1f} MB."
    
    if file_size == 0:
        return False, "File is empty."
    
    return True, None

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    mode = request.form.get("mode")
    file = request.files.get("file")
    
    # Validate mode
    if not mode or mode not in MODES:
        return jsonify({"error": "Invalid conversion mode."}), 400
    
    # Validate file
    allowed_exts = MODES[mode]["input_ext"]
    max_size = app.config["MAX_FILE_SIZE"]
    is_valid, error_msg = validate_file(file, allowed_exts, max_size)
    if not is_valid:
        return jsonify({"error": error_msg}), 400
    
    uid = str(uuid.uuid4())
    src = os.path.join(app.config["UPLOAD_FOLDER"], uid + "_" + file.filename)
    out = os.path.splitext(src)[0] + MODES[mode]["ext"]
    file.save(src)
    try:
        MODES[mode]["fn"](src, out)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(src):
            os.remove(src)
    out_name = os.path.splitext(file.filename)[0] + "_converted" + MODES[mode]["ext"]

    @after_this_request
    def remove_output(response):
        try:
            if os.path.exists(out):
                os.remove(out)
        except Exception:
            pass
        return response

    return send_file(out, as_attachment=True, download_name=out_name)

if __name__ == "__main__":
    debug_mode = os.getenv("FLASK_DEBUG", "False").lower() == "true"
    port = int(os.getenv("FLASK_PORT", 5000))
    host = os.getenv("HOST", "0.0.0.0")
    app.run(debug=debug_mode, host=host, port=port)
