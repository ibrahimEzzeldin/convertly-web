# Convertly-Web

Online PDF toolkit — convert, merge, split, compress, translate, sign, edit, watermark, and more.

## Tech Stack

- **Backend:** Python 3.11 / Flask 3.x
- **PDF engine:** PyMuPDF (fitz), ReportLab, pdfminer, pdf2docx, pdfplumber
- **Frontend:** Single-page app in `templates/index.html` (Jinja2), vanilla JS + Bootstrap
- **Security:** flask-wtf CSRF, Flask-Limiter, custom `security_manager.py` (quota, auth, DDoS)
- **Translation:** MyMemory free API (`translation_service.py`) — NOT Anthropic/Claude
- **Payments:** PayPal (order creation + capture)
- **RTL support:** arabic-reshaper + python-bidi for Arabic/Hebrew PDF rendering
- **Deployment:** Render (Dockerfile + gunicorn), also runs locally on Windows

## Project Structure

```
app.py                  # Main Flask app (~4000 lines, all routes)
security_manager.py     # Quota, auth, fingerprinting, captcha
translation_service.py  # MyMemory translation wrapper
templates/
  index.html            # Main SPA template (~2800 lines)
  base.html, invoice.html, privacy.html, share.html, support.html
static/
  fonts/, favicon.svg, og-image.png, sitemap.xml, manifest.json
uploads/                # Temp file storage (auto-cleaned)
quota.db                # SQLite quota database
```

## Running Locally

```bash
source .venv/Scripts/activate
python app.py
# Runs on http://127.0.0.1:5000
```

Background server:
```bash
source .venv/Scripts/activate && python app.py > server.log 2>&1 &
```

## Key Routes (app.py)

| Route | Method | Purpose |
|---|---|---|
| `/` | GET | Landing page |
| `/convert` | POST | File conversion (PDF/Word/Excel) |
| `/merge-pdf` | POST | Merge multiple PDFs |
| `/split-pdf` | POST | Split PDF by page ranges |
| `/remove-pages` | POST | Remove pages from PDF |
| `/extract-pages` | POST | Extract specific pages |
| `/compress-pdf` | POST | Compress PDF (lossless or rendered) |
| `/organize-pdf/preview` | POST | Get page thumbnails for reorder UI |
| `/organize-pdf/reorder` | POST | Apply page reorder |
| `/pdf-to-jpg` | POST | Convert PDF pages to JPEG |
| `/jpg-to-pdf` | POST | Convert images to PDF |
| `/watermark-pdf` | POST | Add text watermark |
| `/rotate-pdf` | POST | Rotate pages |
| `/unlock-pdf` | POST | Remove PDF password |
| `/protect-pdf` | POST | Add PDF password |
| `/page-numbers` | POST | Add page numbers |
| `/sign-pdf` | POST | Add signature image |
| `/edit-pdf` | POST | Add/edit text on PDF (click-to-place) |
| `/pdf-page-preview` | POST | Get page preview image (no quota) |
| `/pdf-text-extract` | POST | Extract text blocks from a page |
| `/translate-pdf` | POST | Translate PDF text |
| `/translate` | POST | Translate plain text |
| `/create-paypal-order` | POST | PayPal payment flow |
| `/apply-voucher` | POST | Validate voucher code |
| `/redeem-voucher` | POST | Redeem voucher for pro access |
| `/support` | GET/POST | Support contact form |

## Quota System

- 3 free conversions per session (tracked by `ConversionCounter` in `security_manager.py`)
- Pro access via PayPal payment or voucher codes
- `check_conversion_quota` decorator on routes that consume quota
- `/pdf-page-preview` does NOT consume quota

## Development Notes

- **No Tesseract:** OCR routes return 503 — Tesseract is not installed
- **Repair PDF:** Feature removed (no backend route, no frontend UI)
- **File cleanup:** `cleanup_old_files()` runs before requests, removes files older than configured max age
- **CSRF:** All POST forms need `csrf_token()` — handled by flask-wtf
- **Uploads dir:** Files are saved with UUID prefixes, auto-deleted after processing
- **Environment:** Uses `.env` file (loaded with python-dotenv) for secrets (PayPal, JWT, etc.)

## Testing

```bash
source .venv/Scripts/activate
pytest test_features.py
```
