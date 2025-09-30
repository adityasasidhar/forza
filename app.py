from flask import Flask, request, render_template, send_file, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
from tools import read_docx, classify_balance_sheet, extract_json, write_to_pdf

# Use /tmp for Vercel (serverless environment)
UPLOAD_FOLDER = '/tmp'
ALLOWED_EXTENSIONS = {'docx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecretkey')

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            # Process file
            text = read_docx(filepath)
            result_json = classify_balance_sheet(text)
            result_json = extract_json(result_json)
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reclassified_output.pdf')
            write_to_pdf(result_json, pdf_path)
            # After processing, show result page with download link only
            return render_template('result.html')
        else:
            flash('Allowed file type is .docx')
            return redirect(request.url)
    return render_template('upload.html')

@app.route('/download')
def download_pdf():
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reclassified_output.pdf')
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    else:
        flash('PDF not found. Please upload and process a file first.')
        return redirect(url_for('upload_file'))

# Export app for Vercel
app = app