from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
from docx import Document
from PyPDF2 import PdfReader
import mammoth
import pdfplumber

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Required for flash messages

# Configuration
UPLOAD_FOLDER = 'uploads'
HTML_FOLDER = 'html_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(HTML_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Routes
@app.route('/')
def index():
    """Render the homepage with an upload form."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file uploads and conversion."""
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Determine the file type and convert
    html_path = os.path.join(HTML_FOLDER, f"{os.path.splitext(filename)[0]}.html")
    try:
        if filename.endswith('.docx'):
            convert_docx_to_html(file_path, html_path)
        elif filename.endswith('.pdf'):
            convert_pdf_to_html(file_path, html_path)
        else:
            flash('Unsupported file format')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error processing file: {str(e)}')
        return redirect(url_for('index'))

    flash('File converted successfully!')
    return send_file(html_path, as_attachment=True)

# Conversion Functions
def convert_docx_to_html(input_path, output_path):
    """Convert a Word document to HTML using Mammoth."""
    with open(input_path, 'rb') as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # Extract HTML content
    with open(output_path, 'w') as f:
        f.write(html)

def convert_pdf_to_html(input_path, output_path):
    """Convert a PDF to HTML using PDFPlumber."""
    html = """
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; padding: 10px; }
            table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
            td, th { border: 1px solid black; padding: 5px; text-align: left; vertical-align: top; }
            .checkbox { display: inline-block; width: 16px; height: 16px; border: 1px solid black; margin-right: 5px; }
            label { display: inline-block; margin-right: 15px; font-size: 14px; vertical-align: middle; }
        </style>
    </head>
    <body>
    """

    with pdfplumber.open(input_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            html += f"<h2>Page {page_number}</h2>"

            # Extract tables
            tables = page.extract_tables()
            for table in tables:
                html += "<table>"
                for row in table:
                    html += "<tr>"
                    for cell in row:
                        cell = cell or ""
                        # Detect and replace checkbox placeholders
                        if "[ ]" in cell:
                            options = cell.split("[ ]")
                            cell = "".join(f"<label><input type='checkbox' class='checkbox'> {opt.strip()}</label>" for opt in options if opt.strip())
                        html += f"<td>{cell}</td>"
                    html += "</tr>"
                html += "</table>"

            # Extract raw text
            text = page.extract_text()
            if text:
                html += f"<pre>{text}</pre>"

    html += "</body></html>"

    with open(output_path, 'w') as f:
        f.write(html)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
