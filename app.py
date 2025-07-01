# import packages
import os, time, warnings, zipfile, shutil
import pandas as pd
from datetime import datetime
import logging

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename

from invoice_info import get_inv_info
from utils_local.docx_utils import populate_docx_table

warnings.filterwarnings('ignore')

# initialize flask app
app = Flask(__name__)
CORS(app)

# config folders
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['TEMP_FOLDER'] = 'temp'
app.config['TEMPLATE_FOLDER'] = 'templates'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# ensure folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)


# define allowed file types for invoice info file
def allowed_info_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['xlsx', 'xls']


# define allowed file types for invoice template file
def allowed_template_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['docx']


# index route
@app.route('/')
def index():
    return render_template('index.html')


# upload invoice template 
@app.route('/upload/template', methods=['POST'])
def upload_template():
    try:
        # check if file is uploaded
        if 'template' not in request.files:
            return jsonify({'error': 'No template file selected'}), 400

        file = request.files['template']

        # check if file name is empty
        if file.filename == '' or file.filename is None:
            return jsonify({'error': 'No template file selected'}), 400

        # check if file type is allowed
        if not allowed_template_file(file.filename):
            return jsonify({'error': 'Unsupported file type, please upload a Word document (.docx)'}), 400

        # save template file
        filename = secure_filename(file.filename or '')
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
        file.save(template_path)

        return jsonify({
            'message': f'Template uploaded successfully: {filename}',
            'template_name': filename
        })

    except Exception as e:
        return jsonify({'error': f'Error uploading template: {str(e)}'}), 500


# upload invoice info file
@app.route('/upload', methods=['POST'])
def upload_file():
    start_time = time.time()
    try:
        # check if file is uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400

        file = request.files['file']

        # check if file name is empty
        if file.filename == '' or file.filename is None:
            return jsonify({'error': 'No file selected'}), 400

        # check if file type is allowed
        if not allowed_info_file(file.filename):
            return jsonify({'error': 'Unsupported file type, please upload an Excel file (.xlsx or .xls)'}), 400

        # get sales tax rate
        sales_tax_rate = request.form.get('sales_tax_rate', 0.1)  # default 10%
        try:
            sales_tax_rate = float(sales_tax_rate)
            if sales_tax_rate < 0 or sales_tax_rate > 1:
                return jsonify({'error': 'Sales tax rate must be between 0-100%'}), 400
        except ValueError:
            return jsonify({'error': 'Sales tax rate format invalid'}), 400

        # save uploaded file
        filename = secure_filename(file.filename or '')
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # read Excel file
        try:
            inv_info = pd.read_excel(filepath)
        except Exception as e:
            return jsonify({'error': f'Cannot read Excel file: {str(e)}'}), 400

        # check if required columns exist
        required_columns = ['Invoice No.', 'Customer', 'Customer Address1', 'Customer Address2',
                          'Payment Terms', 'Invoice Date', 'Item', 'Detail', 'Unit Price', 'Quantity']

        missing_columns = [col for col in required_columns if col not in inv_info.columns]
        if missing_columns:
            return jsonify({'error': f'"Required columns missing: {", ".join(missing_columns)}'}), 400

        # create temporary folder for storing generated Word documents
        temp_dir = os.path.join(app.config['TEMP_FOLDER'], f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(temp_dir, exist_ok=True)

        generated_files = []

        # generate Word document for each invoice number
        for inv_no in inv_info['Invoice No.'].unique():
            try:
                # get invoice info dictionary, using user specified tax rate
                item_dict = get_inv_info(inv_info, inv_no, sales_tax_rate)

                # generate docx file name
                docx_filename = f"{item_dict['INV_NO']}.docx"
                docx_path = os.path.join(temp_dir, docx_filename)

                # select the template file to use
                custom_template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
                if os.path.exists(custom_template_path):
                    template_file = custom_template_path
                else:
                    template_file = 'assets/template_invoice_format.docx'

                # replace invoice info in Word template
                populate_docx_table(item_dict, template_file, docx_path)

                # check if Word document is generated successfully
                if os.path.exists(docx_path):
                    # add generated Word document to list
                    generated_files.append(docx_filename)
                else:
                    return jsonify({'error': f'Failed to generate Word document: {item_dict["INV_NO"]}'}), 500

            except Exception as e:
                return jsonify({'error': f'Error generating invoice {inv_no}: {str(e)}'}), 500

        # create ZIP file containing all generated files
        zip_filename = f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(app.config['TEMP_FOLDER'], zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in generated_files:
                file_path = os.path.join(temp_dir, file)
                if os.path.exists(file_path):
                    zipf.write(file_path, file)
                else:
                    logging.warning(f"File not found: {file_path}")

        # clean up temporary folder
        shutil.rmtree(temp_dir)

        # delete uploaded Excel file
        os.remove(filepath)

        # calculate total processing time
        total_time = time.time() - start_time

        # determine file types
        file_types = set()
        for file in generated_files:
            if file.endswith('.docx'):
                file_types.add('Word')

        file_type_str = ' & '.join(sorted(file_types)) if file_types else 'Unknown'

        return jsonify({
            'message': f'Invoice generated successfully (Tax rate: {sales_tax_rate*100:.1f}%, Format: {file_type_str})',
            'generated_files': generated_files,
            'zip_file': zip_filename,
            'download_url': f'/download/{zip_filename}',
            'processing_time': f'{total_time:.1f}s',
            'tax_rate': f'{sales_tax_rate*100:.1f}%',
            'file_types': list(file_types)
        })

    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500


# download generated invoices zip file
@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(os.path.join(app.config['TEMP_FOLDER'], filename), as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': f'{filename} not found'}), 404


# download invoice info template
@app.route('/download/template/format')
def download_format_template():
    try:
        return send_file('assets/template_invoice_format.docx', as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': 'Invoice info template not found'}), 404


# download invoice format template
@app.route('/download/template/info')
def download_info_template():
    try:
        return send_file('assets/template_invoice_info.xlsx', as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': 'Invoice format template not found'}), 404


# template status check
@app.route('/template/status')
def template_status():
    try:
        custom_template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
        if os.path.exists(custom_template_path):
            
            return jsonify({
                'has_custom_template': True,
                'template_name': 'custom_template.docx',
                'upload_time': datetime.fromtimestamp(os.stat(custom_template_path).st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'file_size': os.stat(custom_template_path).st_size
            })
        else:
            return jsonify({
                'has_custom_template': False,
                'template_name': 'template_invoice_format.docx'
            })
    except Exception as e:
        return jsonify({'error': f'Error checking template status: {str(e)}'}), 500


# app health check
@app.route('/health')
def health_check():
    return jsonify({'message': 'Service is running normally', 'status': 'healthy'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
