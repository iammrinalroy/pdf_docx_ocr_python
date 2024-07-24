from flask import Flask, redirect, render_template, request, send_file, flash
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import os
from docx import Document
#from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

UPLOAD_FOLDER = 'uploads/'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
pytesseract.pytesseract.tesseract_cmd=r'D:/pdf_docx_ocr_python/Tesseract-OCR/tesseract.exe'

# Function to check if a file has a permitted extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# OCR function using pytesseract
def ocr_core(file, poppler_path):
    images = convert_from_path(file, poppler_path=poppler_path)
    text = ''
    for img in images:
        text += pytesseract.image_to_string(img)
    return text

# Route for home page
@app.route('/')
def home():
    return render_template('index.html')

# Route to handle file upload and OCR process
@app.route('/convert', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = file.filename
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        #Change Poppler path here 
        poppler_path= r'D:/pdf_docx_ocr_python/poppler/Library/bin'

        # Perform OCR
        extracted_text = ocr_core(pdf_path,poppler_path)
        
        # Create DOCX document
        doc = Document()
        doc.add_paragraph(extracted_text)
        
        # Save DOCX
        docx_filename = os.path.splitext(filename)[0] + '.docx'
        doc.save(os.path.join(app.config['UPLOAD_FOLDER'], docx_filename))
        
        return send_file(os.path.join(app.config['UPLOAD_FOLDER'], docx_filename), as_attachment=True)
    else:
        flash('Allowed file type is PDF')
        return redirect(request.url)

if __name__ == '__main__':
    app.run(debug=True)
