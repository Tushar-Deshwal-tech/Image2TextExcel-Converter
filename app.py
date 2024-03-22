from flask import Flask, render_template, request, redirect, url_for, send_file
from PIL import Image
import pytesseract
import tablib
import io

app = Flask(__name__)

# Path to Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Home page
@app.route('/')
def index():
    return render_template('index.html')

# Image to Text conversion route
@app.route('/image_to_text', methods=['POST'])
def image_to_text():
    if 'image' not in request.files:
        return redirect(url_for('index'))

    image_file = request.files['image']

    if image_file.filename == '':
        return redirect(url_for('index'))

    # Convert image to text
    extracted_text = pytesseract.image_to_string(Image.open(image_file))

    return render_template('result.html', text=extracted_text)

# Image to Excel conversion route
@app.route('/image_to_excel', methods=['POST'])
def image_to_excel():
    if 'image' not in request.files:
        return redirect(url_for('index'))

    image_file = request.files['image']

    if image_file.filename == '':
        return redirect(url_for('index'))

    # Convert image to text
    extracted_text = pytesseract.image_to_string(Image.open(image_file))
    
    # Create Excel file with extracted text
    data = tablib.Dataset()
    data.append(['Extracted Text'])
    data.append([extracted_text])

    excel_data = data.export('xlsx')

    # Return the Excel file as response
    return send_file(io.BytesIO(excel_data),
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     attachment_filename='extracted_text.xlsx',
                     as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
