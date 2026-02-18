@"
from flask import Flask, render_template, request, send_file, jsonify
import os, uuid

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs("uploads", exist_ok=True)

def pdf_to_word(src, out):
    from pdf2docx import Converter
    cv = Converter(src); cv.convert(out); cv.close()

def pdf_to_excel(src, out):
    import pdfplumber, openpyxl
    wb = openpyxl.Workbook(); ws = wb.active
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
                    for line in text.split("\n"): ws.append([line])
    wb.save(out)

def word_to_pdf(src, out):
    from docx2pdf import convert
    convert(src, out)

def excel_to_pdf(src, out):
    import subprocess
    subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",os.path.dirname(out),src])

MODES = {
    "pdf-to-word":  {"fn": pdf_to_word,  "ext": ".docx"},
    "pdf-to-excel": {"fn": pdf_to_excel, "ext": ".xlsx"},
    "word-to-pdf":  {"fn": word_to_pdf,  "ext": ".pdf"},
    "excel-to-pdf": {"fn": excel_to_pdf, "ext": ".pdf"},
}

@app.route("/")
def index(): return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    mode = request.form.get("mode")
    file = request.files.get("file")
    if not mode or mode not in MODES: return jsonify({"error": "Invalid mode."}), 400
    if not file or file.filename == "": return jsonify({"error": "No file."}), 400
    uid = str(uuid.uuid4())
    src = os.path.join(app.config["UPLOAD_FOLDER"], uid + "_" + file.filename)
    out = os.path.splitext(src)[0] + MODES[mode]["ext"]
    file.save(src)
    try:
        MODES[mode]["fn"](src, out)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(src): os.remove(src)
    out_name = os.path.splitext(file.filename)[0] + "_converted" + MODES[mode]["ext"]
    return send_file(out, as_attachment=True, download_name=out_name)

if __name__ == "__main__": app.run(debug=True)
"@ | Out-File -FilePath "app.py" -Encoding utf8