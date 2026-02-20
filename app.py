from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from flask_wtf.csrf import CSRFProtect, CSRFError
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from dotenv import load_dotenv
import os, uuid, time, logging, threading
from pathlib import Path

load_dotenv(override=True)

# ── Logging ────────────────────────────────────────────────────────────────
log_level = logging.DEBUG if os.getenv("FLASK_DEBUG", "False").lower() == "true" else logging.INFO
logging.basicConfig(
    level=log_level,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("app.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

# ── App setup ──────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["UPLOAD_FOLDER"]        = os.getenv("UPLOAD_FOLDER", "uploads")
app.config["MAX_CONTENT_LENGTH"]   = int(os.getenv("MAX_CONTENT_LENGTH", 32 * 1024 * 1024))
app.config["MAX_FILE_SIZE"]        = int(os.getenv("MAX_FILE_SIZE", 32 * 1024 * 1024))
app.config["FILE_EXPIRY_HOURS"]    = int(os.getenv("FILE_EXPIRY_HOURS", 24))
app.config["SECRET_KEY"]           = os.getenv("SECRET_KEY", "dev-key-change-in-production")
app.config["CONVERSION_TIMEOUT"]   = int(os.getenv("CONVERSION_TIMEOUT", 120))

session_cookie_secure = os.getenv("SESSION_COOKIE_SECURE")
if session_cookie_secure is None:
    app.config["SESSION_COOKIE_SECURE"] = os.getenv("FLASK_ENV", "production").lower() == "production"
else:
    app.config["SESSION_COOKIE_SECURE"] = session_cookie_secure.lower() == "true"
app.config["SESSION_COOKIE_HTTPONLY"] = os.getenv("SESSION_COOKIE_HTTPONLY", "True").lower() == "true"
app.config["SESSION_COOKIE_SAMESITE"] = os.getenv("SESSION_COOKIE_SAMESITE", "Lax")

# ── CSRF ───────────────────────────────────────────────────────────────────
csrf = CSRFProtect(app)

@app.errorhandler(CSRFError)
def handle_csrf_error(e):
    logger.warning("CSRF validation failed: %s", e.description)
    payload = {"error": "CSRF token validation failed. Please refresh and try again."}
    if app.debug or os.getenv("CSRF_ERROR_DETAILS", "False").lower() == "true":
        payload["details"] = e.description
    return jsonify(payload), 400

# ── Rate limiting ──────────────────────────────────────────────────────────
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri=os.getenv("RATELIMIT_STORAGE_URI", "memory://"),
)

# ── Uploads folder ─────────────────────────────────────────────────────────
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# ── Periodic cleanup ───────────────────────────────────────────────────────
_last_cleanup   = 0.0
_CLEANUP_INTERVAL = 300  # seconds

def cleanup_old_files(max_age_hours=None):
    if max_age_hours is None:
        max_age_hours = app.config["FILE_EXPIRY_HOURS"]
    upload_folder   = Path(app.config["UPLOAD_FOLDER"])
    if not upload_folder.exists():
        return
    current_time    = time.time()
    max_age_seconds = max_age_hours * 3600
    try:
        for file_path in upload_folder.glob("*"):
            if file_path.is_file() and (current_time - file_path.stat().st_mtime) > max_age_seconds:
                file_path.unlink()
                logger.info("Cleaned up old file: %s", file_path.name)
    except Exception as exc:
        logger.error("Error during file cleanup: %s", exc)

@app.before_request
def cleanup_before_request():
    global _last_cleanup
    now = time.time()
    if now - _last_cleanup >= _CLEANUP_INTERVAL:
        _last_cleanup = now
        cleanup_old_files()

# ── Conversion helpers ─────────────────────────────────────────────────────

def _run_with_timeout(fn, args, timeout_seconds):
    """Run fn(*args) in a thread; raise TimeoutError if it exceeds timeout."""
    result    = [None]
    exception = [None]

    def target():
        try:
            result[0] = fn(*args)
        except Exception as exc:
            exception[0] = exc

    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join(timeout_seconds)

    if t.is_alive():
        raise TimeoutError(f"Conversion exceeded {timeout_seconds}s time limit.")
    if exception[0]:
        raise exception[0]
    return result[0]


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
    """
    Convert DOCX → PDF.
    Primary:  docx2pdf  (requires Microsoft Word on Windows / LibreOffice on Linux)
    Fallback: mammoth (DOCX→HTML) + reportlab (HTML text → PDF) — works everywhere.
    """
    try:
        from docx2pdf import convert
        convert(src, out)
        logger.info("word_to_pdf: used docx2pdf")
    except Exception as primary_err:
        logger.warning("word_to_pdf docx2pdf failed (%s), falling back to mammoth", primary_err)
        _word_to_pdf_fallback(src, out)


def _word_to_pdf_fallback(src, out):
    """Pure-Python DOCX → PDF via mammoth (HTML extraction) + reportlab."""
    import mammoth
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.enums import TA_LEFT
    import html, re

    with open(src, "rb") as f:
        result = mammoth.convert_to_html(f)
    raw_html = result.value

    # Strip HTML tags; split into paragraphs on block elements
    raw_html = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.IGNORECASE)
    raw_html = re.sub(r"</?(p|div|li|h[1-6])[^>]*>", "\n", raw_html, flags=re.IGNORECASE)
    raw_html = re.sub(r"<[^>]+>", "", raw_html)
    text     = html.unescape(raw_html)

    doc      = SimpleDocTemplate(out, pagesize=A4,
                                  topMargin=0.75 * inch, bottomMargin=0.75 * inch,
                                  leftMargin=inch, rightMargin=inch)
    styles   = getSampleStyleSheet()
    body     = ParagraphStyle("body", parent=styles["Normal"], fontSize=10,
                               leading=14, spaceAfter=4, alignment=TA_LEFT)
    elements = []
    for line in text.split("\n"):
        line = line.strip()
        if line:
            elements.append(Paragraph(line, body))
        else:
            elements.append(Spacer(1, 6))
    doc.build(elements)
    logger.info("word_to_pdf: used mammoth fallback")


def excel_to_pdf(src, out):
    import openpyxl
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors

    wb = openpyxl.load_workbook(src)
    ws = wb.active

    data = []
    for row in ws.iter_rows(values_only=True):
        data.append([str(cell) if cell is not None else "" for cell in row])

    if not data:
        # Empty sheet — write a blank PDF
        doc = SimpleDocTemplate(out, pagesize=A4)
        doc.build([])
        return

    col_count  = len(data[0])
    # Use landscape for wide sheets
    page_size  = landscape(A4) if col_count > 6 else A4
    page_width = page_size[0] - 1.0 * inch
    col_width  = min(1.5 * inch, page_width / col_count) if col_count else 1.5 * inch

    doc   = SimpleDocTemplate(out, pagesize=page_size,
                               topMargin=0.5 * inch, bottomMargin=0.5 * inch,
                               leftMargin=0.5 * inch, rightMargin=0.5 * inch)
    table = Table(data, colWidths=[col_width] * col_count, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#4f6ef7")),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("FONTSIZE",      (0, 1), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  10),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f2ff")]),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#c5caff")),
    ]))
    doc.build([table])


# ── Conversion modes ───────────────────────────────────────────────────────
MODES = {
    "pdf-to-word":  {"fn": pdf_to_word,  "ext": ".docx", "input_ext": [".pdf"]},
    "pdf-to-excel": {"fn": pdf_to_excel, "ext": ".xlsx", "input_ext": [".pdf"]},
    "word-to-pdf":  {"fn": word_to_pdf,  "ext": ".pdf",  "input_ext": [".docx", ".doc"]},
    "excel-to-pdf": {"fn": excel_to_pdf, "ext": ".pdf",  "input_ext": [".xlsx", ".xls"]},
}

ALLOWED_MIME_TYPES = {
    ".pdf":  ["application/pdf"],
    ".docx": ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    ".doc":  ["application/msword"],
    ".xlsx": ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    ".xls":  ["application/vnd.ms-excel"],
}

# ── File validation ────────────────────────────────────────────────────────
def validate_file(file, allowed_extensions, max_size):
    if not file or file.filename == "":
        return False, "No file provided."

    file_ext = os.path.splitext(file.filename)[1].lower()

    if file_ext not in allowed_extensions:
        return False, f"Invalid file type. Allowed: {', '.join(allowed_extensions)}"

    if hasattr(file, "content_type"):
        allowed_mimes = ALLOWED_MIME_TYPES.get(file_ext, [])
        if allowed_mimes and file.content_type not in allowed_mimes:
            return False, f"Invalid file format for {file_ext} file."

    file.seek(0, 2)
    file_size = file.tell()
    file.seek(0)

    if file_size > max_size:
        return False, f"File too large. Maximum size is {max_size / (1024*1024):.0f} MB."
    if file_size == 0:
        return False, "File is empty."

    return True, None

# ── Routes ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
@limiter.limit(os.getenv("CONVERT_RATE_LIMIT", "10 per minute"))
def convert():
    mode = request.form.get("mode")
    file = request.files.get("file")

    if not mode or mode not in MODES:
        return jsonify({"error": "Invalid conversion mode."}), 400

    allowed_exts = MODES[mode]["input_ext"]
    is_valid, error_msg = validate_file(file, allowed_exts, app.config["MAX_FILE_SIZE"])
    if not is_valid:
        return jsonify({"error": error_msg}), 400

    uid = str(uuid.uuid4())
    # Sanitise filename — strip path components
    safe_name = os.path.basename(file.filename)
    src = os.path.join(app.config["UPLOAD_FOLDER"], f"{uid}_{safe_name}")
    out = os.path.splitext(src)[0] + MODES[mode]["ext"]
    file.save(src)
    logger.info("Converting [%s] %s", mode, safe_name)

    try:
        _run_with_timeout(
            MODES[mode]["fn"],
            (src, out),
            app.config["CONVERSION_TIMEOUT"],
        )
    except TimeoutError as exc:
        logger.error("Conversion timeout for %s: %s", safe_name, exc)
        return jsonify({"error": str(exc)}), 504
    except Exception as exc:
        logger.error("Conversion error for %s: %s", safe_name, exc, exc_info=True)
        return jsonify({"error": "Conversion failed. Please check your file and try again."}), 500
    finally:
        if os.path.exists(src):
            os.remove(src)

    if not os.path.exists(out):
        logger.error("Output file missing after conversion: %s", out)
        return jsonify({"error": "Conversion produced no output. Please try again."}), 500

    out_name = os.path.splitext(safe_name)[0] + "_converted" + MODES[mode]["ext"]
    logger.info("Conversion complete: %s → %s", safe_name, out_name)

    @after_this_request
    def remove_output(response):
        try:
            if os.path.exists(out):
                os.remove(out)
        except Exception:
            pass
        return response

    return send_file(out, as_attachment=True, download_name=out_name)


@app.errorhandler(429)
def ratelimit_error(e):
    logger.warning("Rate limit exceeded from %s", get_remote_address())
    return jsonify({"error": "Too many requests. Please wait a moment and try again."}), 429


@app.errorhandler(413)
def file_too_large(e):
    return jsonify({"error": f"File too large. Maximum size is {app.config['MAX_FILE_SIZE'] // (1024*1024)} MB."}), 413


if __name__ == "__main__":
    debug_mode = os.getenv("FLASK_DEBUG", "False").lower() == "true"
    port       = int(os.getenv("FLASK_PORT", 5000))
    host       = os.getenv("HOST", "0.0.0.0")
    app.run(debug=debug_mode, host=host, port=port)
