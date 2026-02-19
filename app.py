from flask import Flask, render_template, request, send_file, jsonify
import os, uuid

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

os.makedirs("uploads", exist_ok=True)

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
    from reportlab.lib.pagesizes import letter, A4
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
        # Create table
        table = Table(data, colWidths=[1.5*inch]*min(len(data[0]), 5))
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
    "pdf-to-word":  {"fn": pdf_to_word,  "ext": ".docx"},
    "pdf-to-excel": {"fn": pdf_to_excel, "ext": ".xlsx"},
    "word-to-pdf":  {"fn": word_to_pdf,  "ext": ".pdf"},
    "excel-to-pdf": {"fn": excel_to_pdf, "ext": ".pdf"},
}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    mode = request.form.get("mode")
    file = request.files.get("file")
    if not mode or mode not in MODES:
        return jsonify({"error": "Invalid mode."}), 400
    if not file or file.filename == "":
        return jsonify({"error": "No file."}), 400
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
    return send_file(out, as_attachment=True, download_name=out_name)

if __name__ == "__main__":
    app.run(debug=True)