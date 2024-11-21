import os
import uuid
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
import PyPDF2
import tabula
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_tables_with_password(file_path, password=None):
    """
    Extract tables from PDF, handling password-protected PDFs
    
    Args:
        file_path (str): Path to the PDF file
        password (str, optional): Password for encrypted PDF
    
    Returns:
        list: Extracted tables as pandas DataFrames
    """
    try:
        # Try extracting tables using tabula
        tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
        
        # Filter out small or empty tables
        tables = [table for table in tables if not table.empty and len(table.columns) > 1]
        
        return tables
    
    except Exception as e:
        # Handle specific PDF-related exceptions
        if "cannot decrypt" in str(e).lower():
            raise ValueError("PDF is encrypted. Password required.")
        
        return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    password = request.form.get('password', None)
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4()}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        
        try:
            tables = extract_tables_with_password(file_path, password)
            
            if not tables:
                return jsonify({"error": "No tables found in the PDF"}), 404
            
            # Select the largest table (assuming it's the most relevant)
            main_table = max(tables, key=len)
            
            # Save table to Excel
            excel_filename = f"{uuid.uuid4()}_extracted_table.xlsx"
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            main_table.to_excel(excel_path, index=False)
            
            return jsonify({
                "table_data": main_table.to_dict(orient='records'),
                "excel_filename": excel_filename
            }), 200
        
        except ValueError as ve:
            return jsonify({"error": str(ve)}), 403
        
        except Exception as e:
            return jsonify({"error": f"Error processing file: {str(e)}"}), 500
    
    return jsonify({"error": "Invalid file type"}), 400

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(
            os.path.join(app.config['UPLOAD_FOLDER'], filename),
            as_attachment=True
        )
    except Exception as e:
        return str(e), 404

if __name__ == '__main__':
    app.run(debug=True)
