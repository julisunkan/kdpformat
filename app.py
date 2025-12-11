"""
KDP Book Formatter Web Application
A Flask-based tool for formatting manuscripts according to Amazon KDP standards.
"""

import os
import uuid
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
from werkzeug.utils import secure_filename

from services.formatter import format_manuscript
from services.pdf_exporter import convert_to_pdf, is_pdf_conversion_available

app = Flask(__name__)
app.secret_key = os.environ.get('SESSION_SECRET', 'dev-secret-key-change-in-production')

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Display the upload page."""
    return render_template('upload.html')


@app.route('/process', methods=['POST'])
def process():
    """Process the uploaded manuscript."""
    if 'file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a .docx file', 'error')
        return redirect(url_for('index'))
    
    trim_size = request.form.get('trim_size', '6x9')
    print_mode = request.form.get('print_mode') == 'on'
    title = request.form.get('title', 'Untitled')
    author = request.form.get('author', 'Author Name')
    line_spacing = float(request.form.get('line_spacing', '1.15'))
    generate_pdf = request.form.get('generate_pdf') == 'on'
    
    unique_id = uuid.uuid4().hex[:8]
    original_filename = secure_filename(file.filename or 'document.docx')
    base_name = os.path.splitext(original_filename)[0]
    
    input_filename = f"input_{unique_id}_{original_filename}"
    output_filename = f"KDP_FORMATTED_{base_name}_{unique_id}.docx"
    
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    file.save(input_path)
    
    result = format_manuscript(
        input_path=input_path,
        output_path=output_path,
        trim_size=trim_size,
        print_mode=print_mode,
        title=title,
        author=author,
        line_spacing=line_spacing
    )
    
    if not result['success']:
        flash(f"Formatting failed: {result.get('error', 'Unknown error')}", 'error')
        return redirect(url_for('index'))
    
    pdf_path = None
    pdf_error = None
    
    if print_mode and generate_pdf:
        pdf_success, pdf_result = convert_to_pdf(output_path, app.config['UPLOAD_FOLDER'])
        if pdf_success:
            pdf_path = pdf_result
        else:
            pdf_error = pdf_result
    
    session['result'] = {
        'docx_path': output_path,
        'docx_filename': output_filename,
        'pdf_path': pdf_path,
        'pdf_filename': os.path.basename(pdf_path) if pdf_path else None,
        'pdf_error': pdf_error,
        'dpi_warnings': result.get('dpi_warnings', []),
        'title': title,
        'author': author,
        'trim_size': trim_size,
        'print_mode': print_mode
    }
    
    try:
        os.remove(input_path)
    except:
        pass
    
    return redirect(url_for('results'))


@app.route('/results')
def results():
    """Display processing results."""
    result = session.get('result')
    if not result:
        flash('No processing results found', 'error')
        return redirect(url_for('index'))
    
    return render_template('results.html', result=result)


@app.route('/download/<file_type>')
def download(file_type):
    """Download the processed file."""
    result = session.get('result')
    if not result:
        flash('No file available for download', 'error')
        return redirect(url_for('index'))
    
    if file_type == 'docx':
        file_path = result.get('docx_path')
        filename = result.get('docx_filename')
    elif file_type == 'pdf':
        file_path = result.get('pdf_path')
        filename = result.get('pdf_filename')
    else:
        flash('Invalid file type', 'error')
        return redirect(url_for('results'))
    
    if not file_path or not os.path.exists(file_path):
        flash('File not found', 'error')
        return redirect(url_for('results'))
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name=filename
    )


@app.after_request
def add_cache_headers(response):
    """Add cache control headers."""
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
