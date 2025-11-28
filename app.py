"""
Flask Web Application for PDF Invoice Extractor
Allows users to upload PDFs and generate Excel files with configurable accounting columns.
"""

from flask import Flask, render_template, request, send_file, jsonify, session
import os
from werkzeug.utils import secure_filename
import uuid
from pdf_extractor import extract_invoices
from excel_generator import generate_excel

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle PDF file upload and extraction."""
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['pdf_file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4())
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
        file.save(upload_path)
        
        # Extract invoice data
        invoices = extract_invoices(upload_path)
        
        # Store session data
        session['upload_path'] = upload_path
        session['unique_id'] = unique_id
        session['invoice_count'] = len(invoices)
        
        # Return extracted data for preview
        return jsonify({
            'success': True,
            'invoice_count': len(invoices),
            'invoices': invoices,
            'unique_id': unique_id
        })
    
    except Exception as e:
        return jsonify({'error': f'Error processing PDF: {str(e)}'}), 500


@app.route('/generate_excel', methods=['POST'])
def generate_excel_file():
    """Generate Excel file with user-provided configuration."""
    try:
        # Get configuration from form
        config = {
            'SATZART': request.form.get('satzart', 'D'),
            'FIRMA': request.form.get('firma', ''),
            'SOLL_HABEN': request.form.get('soll_haben', ''),
            'BUCH_KREIS': request.form.get('buch_kreis', ''),
            'HABENKONTO': request.form.get('habenkonto', ''),
            'KOSTSTELLE': request.form.get('koststelle', ''),
            'KOSTTRAGER': request.form.get('kosttrager', ''),
            'Kostenträgerbezeichnung': request.form.get('kosttraegerbezeichnung', ''),
            'Bebuchbar': request.form.get('bebuchbar', 'Ja'),
            'BUCH_TEXT_PREFIX': request.form.get('buch_text_prefix', '')
        }
        
        # Get upload path from session
        upload_path = session.get('upload_path')
        unique_id = session.get('unique_id')
        
        if not upload_path or not os.path.exists(upload_path):
            return jsonify({'error': 'No uploaded file found. Please upload a PDF first.'}), 400
        
        # Extract invoices again
        invoices = extract_invoices(upload_path)
        
        # Generate Excel file
        output_filename = f"{unique_id}_output.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        generate_excel(invoices, output_path, config)
        
        # Store output path in session
        session['output_path'] = output_path
        
        return jsonify({
            'success': True,
            'output_filename': output_filename,
            'download_url': f'/download/{output_filename}'
        })
    
    except Exception as e:
        return jsonify({'error': f'Error generating Excel: {str(e)}'}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download the generated Excel file."""
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    
    if not os.path.exists(file_path):
        return "File not found", 404
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name=f"invoice_export_{filename}"
    )


if __name__ == '__main__':
    print("Starting PDF Invoice Extractor Web Application...")
    print("Open your browser and navigate to: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
