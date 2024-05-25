from flask import Flask, request, jsonify, render_template_string
from PIL import Image
import pytesseract
import openpyxl
from openpyxl import Workbook
import re
import os
import base64
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
EXCEL_FILE_PATH = 'ocr_data.xlsx'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extract_data(text):
    net_weight = extract_net_weight(text)
    manufacturing_date = extract_manufacturing_date(text)
    mrp = extract_mrp(text)
    batch_number = extract_batch_number(text)

    return {
        "manufacturing_date": manufacturing_date,
        "batch_number": batch_number,
        "net_weight": net_weight,
        "mrp": mrp
    }

def extract_net_weight(text):
    net_weight_regex = re.compile(r'\b\d+g\b', re.IGNORECASE)
    match = net_weight_regex.search(text)
    return match.group(0) if match else 'Not Found'

def extract_manufacturing_date(text):
    date_regex = re.compile(r'\b(\d{2}/\d{2}/\d{2})\b')
    match = date_regex.search(text)
    return match.group(0) if match else 'Not Found'

def extract_mrp(text):
    # Search for patterns like "__.00" or "__,00"
    mrp_regex = re.compile(r'\b\d+[.,]\d{2}\b')
    match = mrp_regex.search(text)
    if match:
        return match.group(0)

    # Search for patterns like "MRP: __"
    mrp_colon_regex = re.compile(r'\bMRP\s*:\s*(\d+[.,]\d{2})\b', re.IGNORECASE)
    match = mrp_colon_regex.search(text)
    if match:
        return match.group(1)

    return 'Not Found'

def extract_batch_number(text):
    batch_number_regex = re.compile(r'\b[A-Z0-9]{7}\b')
    match = batch_number_regex.search(text)
    return match.group(0) if match else 'Not Found'


def save_to_excel(data, excel_file_path):
    try:
        if os.path.exists(excel_file_path):
            workbook = openpyxl.load_workbook(excel_file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(['Manufacturing Date', 'Batch Number', 'Net Weight', 'MRP'])

        sheet.append([data['manufacturing_date'], data['batch_number'], data['net_weight'], data['mrp']])
        workbook.save(excel_file_path)

    except Exception as e:
        print(f'Error saving data to Excel file: {e}')

@app.route('/', methods=['GET', 'POST'])
def index():
    extracted_data = None
    raw_text = None
    if request.method == 'POST':
        if 'file' in request.files:
            file = request.files['file']
            if file.filename == '':
                return jsonify({"error": "No selected file"}), 400

            if file:
                file_path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(file_path)

                # Perform OCR using pytesseract
                img_obj = Image.open(file_path)
                raw_text = pytesseract.image_to_string(img_obj)

                # Extract data from OCR text
                extracted_data = extract_data(raw_text)
                print('Extracted Data:', extracted_data)

                # Save extracted data to Excel
                save_to_excel(extracted_data, EXCEL_FILE_PATH)

        if 'image_data' in request.form:
            image_data = request.form['image_data']
            image_data = re.sub('^data:image/.+;base64,', '', image_data)
            image = Image.open(BytesIO(base64.b64decode(image_data)))
            raw_text = pytesseract.image_to_string(image)

            # Extract data from OCR text
            extracted_data = extract_data(raw_text)
            print('Extracted Data:', extracted_data)

            # Save extracted data to Excel
            save_to_excel(extracted_data, EXCEL_FILE_PATH)
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Upload Image for OCR</title>
       <link rel="stylesheet" type="text/css" href="{{ url_for('static', filename='style.css') }}">
        
    </head>
    <div class="main">
   
    <div class="upload">
      <h1 class="heading">Upload Image for OCR</h1>
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="Upload">
    </form>
    </div>
    <div class="vedioCam"> <h1>Capture Image from Webcam</h1>
    <video id="video" width="640" height="480" autoplay></video>
    <button id="snap">Capture</button>
    <canvas id="canvas" width="640" height="480" style="display:none;"></canvas>
    <form method="post" enctype="multipart/form-data">
        <input type="hidden" name="image_data" id="image_data">
        <input type="submit" value="Upload Captured Image">
    </form></div>
      </div>
    {% if raw_text %}
    <div class="data">
     <div>                             
    <h2>Raw Extracted Text</h2>
    <pre>{{ raw_text }}</pre>
    {% endif %}
    </div>
    <div>                              
    {% if extracted_data %}
    <h2>Extracted Data</h2>
    <ul>
      <li><strong>Manufacturing Date:</strong> {{ extracted_data.manufacturing_date }}</li>
      <li><strong>Batch Number:</strong> {{ extracted_data.batch_number }}</li>
      <li><strong>Net Weight:</strong> {{ extracted_data.net_weight }}</li>
      <li><strong>MRP:</strong> {{ extracted_data.mrp }}</li>
    </ul>
     
    {% endif %}
    </div>
    </div>
    <script>
      const video = document.getElementById('video');
      const canvas = document.getElementById('canvas');
      const context = canvas.getContext('2d');
      const snap = document.getElementById('snap');
      const image_data_input = document.getElementById('image_data');

      navigator.mediaDevices.getUserMedia({ video: true })
        .then((stream) => {
          video.srcObject = stream;
        })
        .catch((err) => {
          console.error("Error accessing the camera: " + err);
        });

      snap.addEventListener('click', () => {
        context.drawImage(video, 0, 0, 640, 480);
        const imageData = canvas.toDataURL('image/png');
        image_data_input.value = imageData;
      });
    </script>
    ''', raw_text=raw_text, extracted_data=extracted_data)

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
