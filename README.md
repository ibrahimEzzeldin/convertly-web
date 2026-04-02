# Convertly - Free Online PDF & Document Toolkit

All-in-one PDF toolkit — convert, merge, split, compress, edit, sign, translate, and more. No sign-up, no installation. Just upload and go.

**Website**: [https://convertly-web.onrender.com](https://convertly-web.onrender.com)

## Tools

### Convert
| Tool | Description |
|---|---|
| PDF to Word | Convert PDF files to editable Word documents (.docx) |
| Word to PDF | Convert Word documents (.docx, .doc) to PDF |
| PDF to Excel | Extract tables from PDFs into Excel spreadsheets (.xlsx) |
| Excel to PDF | Convert Excel spreadsheets to PDF |
| PDF to JPG | Convert PDF pages to high-quality JPEG images |
| JPG to PDF | Combine one or more images (JPG, PNG, WEBP, BMP, GIF, TIFF) into a single PDF |

### Organise
| Tool | Description |
|---|---|
| Merge PDF | Combine multiple PDF files into one (up to 20 files) |
| Split PDF | Split a PDF by page ranges into separate files |
| Remove Pages | Delete specific pages from a PDF |
| Extract Pages | Keep only the pages you need |
| Compress PDF | Reduce PDF file size (lossless, balanced, or maximum compression) |
| Organize PDF | Drag-and-drop page reordering with visual thumbnails |
| Rotate PDF | Rotate pages 90, 180, or 270 degrees |
| Add Page Numbers | Stamp page numbers on every page with custom positioning |

### Edit & Sign
| Tool | Description |
|---|---|
| Edit PDF | Click-to-place text editor with live page preview |
| Watermark PDF | Add text watermarks across all pages |
| Sign PDF | Draw, type, or upload a signature onto your PDF |
| Translate PDF | AI-powered PDF translation supporting 23+ languages |

### Protect
| Tool | Description |
|---|---|
| Protect PDF | Add password protection to your PDF |
| Unlock PDF | Remove password protection from encrypted PDFs |

## Features

- **20+ PDF tools** in one place
- **No sign-up required** — start using immediately
- **Free tier** — 3 free conversions, no registration needed
- **One-time payment** — $2 for 20 additional conversions via PayPal (no subscription)
- **Voucher codes** — promotional codes for free conversions
- **RTL language support** — full Arabic and Hebrew text rendering in PDFs
- **AI translation** — translate PDFs into 23+ languages
- **Mobile-friendly** — responsive design works on any device
- **Files auto-deleted** — uploaded files are removed after processing

## Tech Stack

- **Backend**: Python 3.11 / Flask
- **PDF Engine**: PyMuPDF (fitz), ReportLab, pdfminer, pdf2docx
- **Frontend**: Vanilla JS + Bootstrap (single-page app)
- **Payments**: PayPal
- **Translation**: MyMemory API
- **Deployment**: Docker + Gunicorn on Render

## Getting Started

Visit **[convertly-web.onrender.com](https://convertly-web.onrender.com)** to start using Convertly right away.

### Run Locally

```bash
# Clone the repo
git clone https://github.com/ibrahimEzzeldin/convertly-web.git
cd convertly-web

# Create virtual environment
python -m venv .venv
source .venv/Scripts/activate   # Windows
# source .venv/bin/activate     # Linux/Mac

# Install dependencies
pip install -r requirements.txt

# Set up environment variables
cp .env.example .env
# Edit .env with your keys

# Run the server
python app.py
# Open http://127.0.0.1:5000
```

## Security

- **CSRF Protection** — all forms protected with flask-wtf tokens
- **Content Security Policy** — strict CSP headers on all responses
- **Rate Limiting** — per-endpoint rate limits to prevent abuse
- **File Validation** — file type, MIME type, magic bytes, and size checks
- **Secure Sessions** — HttpOnly, Secure, SameSite cookies
- **Server-side Quota** — fingerprint-based tracking that can't be bypassed client-side
- **Input Sanitization** — log sanitization to prevent credential leaks

## Pricing

| Plan | Conversions | Price |
|---|---|---|
| Free | 3 | $0 |
| Pro | 20 | $2 (one-time) |

No subscription. No account required.

## Contact & Feedback

- Open an issue: [GitHub Issues](https://github.com/ibrahimEzzeldin/convertly-web/issues)
- Support page: [convertly-web.onrender.com/support](https://convertly-web.onrender.com/support)
