from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from docx import Document
from PyPDF2 import PdfReader
from lxml import etree

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
HTML_FOLDER = 'html_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(HTML_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    # Convert file to HTML
    html_path = os.path.join(HTML_FOLDER, f"{os.path.splitext(filename)[0]}.html")
    if filename.endswith('.docx'):
        convert_docx_to_html(file_path, html_path)
    elif filename.endswith('.pdf'):
        convert_pdf_to_html(file_path, html_path)
    else:
        return "Unsupported file format"
    
    return send_file(html_path, as_attachment=True)

import mammoth

def convert_docx_to_html(input_path, output_path):
    with open(input_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # Extract HTML content
    with open(output_path, "w") as f:
        f.write(html)


from pdf2image import convert_from_path
import pytesseract
from lxml import etree

import pdfplumber
from bs4 import BeautifulSoup

def is_duplicate_line(line, seen_lines):
    if line in seen_lines:
        return True
    seen_lines.add(line)
    return False

def convert_pdf_to_html(input_path, output_path):
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
    seen_lines = set()  # Track lines to prevent duplicates

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

            # Extract raw text, filter duplicates and unwanted lines
            text = page.extract_text()
            if text:
                filtered_lines = []
                for line in text.split("\n"):
                    # Skip duplicate lines
                    if line in seen_lines:
                        continue
                    seen_lines.add(line)

                    # Skip specific unwanted lines
                    if "Disclaimer" in line or line.startswith("For Office Use Only"):
                        continue

                    # Append valid lines
                    filtered_lines.append(line)

                # Add filtered text to HTML
                if filtered_lines:
                    html += f"<pre>{'\n'.join(filtered_lines)}</pre>"

    html += "</body></html>"

    # Write the HTML output
    with open(output_path, "w") as f:
        f.write(html)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
