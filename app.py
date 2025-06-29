# import packages
import os, re, time, warnings, zipfile, tempfile, shutil
import pandas as pd
from datetime import datetime
import logging

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename

from invoice_info import get_inv_info
from utils.docx_manipulate import populate_docx_table, convert_docx_pdf

warnings.filterwarnings('ignore')

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))


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

# define allowed file types
def allowed_info_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['xlsx', 'xls']

def allowed_template_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['docx']

# define routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload/template', methods=['POST'])
def upload_template():
    """上传自定义发票模板"""
    try:
        # 检查是否有文件
        if 'template' not in request.files:
            return jsonify({'error': '没有选择模板文件'}), 400

        file = request.files['template']

        # 检查文件名
        if file.filename == '' or file.filename is None:
            return jsonify({'error': '没有选择模板文件'}), 400

        # 检查文件类型
        if not allowed_template_file(file.filename):
            return jsonify({'error': '不支持的文件类型，请上传Word文档(.docx)'}), 400

        # 保存模板文件
        filename = secure_filename(file.filename or '')
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
        file.save(template_path)

        return jsonify({
            'message': f'模板上传成功: {filename}',
            'template_name': filename
        })

    except Exception as e:
        return jsonify({'error': f'上传模板时出错: {str(e)}'}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理Excel文件上传和Invoice生成"""
    start_time = time.time()
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400

        file = request.files['file']

        # 检查文件名
        if file.filename == '' or file.filename is None:
            return jsonify({'error': '没有选择文件'}), 400

        # 检查文件类型
        if not allowed_info_file(file.filename):
            return jsonify({'error': '不支持的文件类型，请上传Excel文件(.xlsx或.xls)'}), 400

        # 获取销售税率
        sales_tax_rate = request.form.get('sales_tax_rate', 0.1)  # 默认10%
        try:
            sales_tax_rate = float(sales_tax_rate)
            if sales_tax_rate < 0 or sales_tax_rate > 1:
                return jsonify({'error': '销售税率必须在0-100%之间'}), 400
        except ValueError:
            return jsonify({'error': '销售税率格式无效'}), 400

        # 保存上传的文件
        filename = secure_filename(file.filename or '')
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # 读取Excel文件
        try:
            inv_info = pd.read_excel(filepath)
        except Exception as e:
            return jsonify({'error': f'无法读取Excel文件: {str(e)}'}), 400

        # 检查必要的列是否存在
        required_columns = ['Invoice No.', 'Customer', 'Customer Address1', 'Customer Address2',
                          'Payment Terms', 'Invoice Date', 'Item', 'Detail', 'Unit Price', 'Quantity']

        missing_columns = [col for col in required_columns if col not in inv_info.columns]
        if missing_columns:
            return jsonify({'error': f'Excel文件缺少必要的列: {", ".join(missing_columns)}'}), 400

        # 创建临时文件夹用于存储生成的PDF
        temp_dir = os.path.join(app.config['TEMP_FOLDER'], f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(temp_dir, exist_ok=True)

        generated_files = []

        # 为每个发票号生成PDF
        for inv_no in inv_info['Invoice No.'].unique():
            try:
                # 获取发票信息字典，使用用户指定的税率
                item_dict = get_inv_info(inv_info, inv_no, sales_tax_rate)

                # 生成docx文件名
                docx_filename = f"{item_dict['INV_NO']}.docx"
                pdf_filename = f"{item_dict['INV_NO']}.pdf"
                docx_path = os.path.join(temp_dir, docx_filename)
                pdf_path = os.path.join(temp_dir, pdf_filename)

                # 选择使用的模板文件
                custom_template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
                if os.path.exists(custom_template_path):
                    template_file = custom_template_path
                else:
                    template_file = 'assets/template_invoice_format.docx'

                # 替换Word模板中的发票信息
                populate_docx_table(item_dict, template_file, docx_path)

                # 直接调用 convert_docx_pdf 进行PDF转换
                convert_docx_pdf(docx_path, keep=False)

                # 检查PDF是否生成
                if os.path.exists(pdf_path):
                    # 删除DOCX文件
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
                    # 添加生成的PDF到列表
                    generated_files.append(pdf_filename)
                else:
                    # 查找生成的PDF（docx2pdf 可能生成的文件名不一致）
                    pdf_dir = os.path.dirname(pdf_path)
                    for file in os.listdir(pdf_dir):
                        if file.endswith('.pdf') and os.path.splitext(os.path.basename(docx_path))[0] in file:
                            actual_pdf = os.path.join(pdf_dir, file)
                            shutil.move(actual_pdf, pdf_path)
                            generated_files.append(pdf_filename)
                            break
                    else:
                        return jsonify({'error': f'转换PDF失败: {item_dict["INV_NO"]}'}), 500

            except Exception as e:
                return jsonify({'error': f'生成发票 {inv_no} 时出错: {str(e)}'}), 500

        # 创建ZIP文件包含所有生成的文件
        zip_filename = f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(app.config['TEMP_FOLDER'], zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in generated_files:
                file_path = os.path.join(temp_dir, file)
                if os.path.exists(file_path):
                    zipf.write(file_path, file)
                else:
                    logging.warning(f"File not found: {file_path}")

        # 清理临时文件夹
        shutil.rmtree(temp_dir)

        # 删除上传的Excel文件
        os.remove(filepath)

        # 计算总处理时间
        total_time = time.time() - start_time

        # 确定文件类型
        file_types = set()
        for file in generated_files:
            if file.endswith('.pdf'):
                file_types.add('PDF')
            elif file.endswith('.docx'):
                file_types.add('DOCX')

        file_type_str = ' & '.join(sorted(file_types)) if file_types else 'Unknown'

        return jsonify({
            'message': f'发票生成成功 (税率: {sales_tax_rate*100:.1f}%, 格式: {file_type_str})',
            'generated_files': generated_files,
            'zip_file': zip_filename,
            'download_url': f'/download/{zip_filename}',
            'processing_time': f'{total_time:.1f}秒',
            'tax_rate': f'{sales_tax_rate*100:.1f}%',
            'file_types': list(file_types)
        })

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载生成的ZIP文件"""
    try:
        return send_file(os.path.join(app.config['TEMP_FOLDER'], filename), as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': '文件不存在'}), 404

@app.route('/download/template/format')
def download_format_template():
    """下载发票格式模板"""
    try:
        return send_file('assets/template_invoice_format.docx', as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': '模板文件不存在'}), 404

@app.route('/download/template/info')
def download_info_template():
    """下载发票信息模板"""
    try:
        return send_file('assets/template_invoice_info.xlsx', as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': '模板文件不存在'}), 404

@app.route('/template/status')
def template_status():
    """检查当前模板状态"""
    try:
        custom_template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'custom_template.docx')
        if os.path.exists(custom_template_path):
            # 获取文件信息
            stat = os.stat(custom_template_path)
            return jsonify({
                'has_custom_template': True,
                'template_name': 'custom_template.docx',
                'upload_time': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'file_size': stat.st_size
            })
        else:
            return jsonify({
                'has_custom_template': False,
                'template_name': 'template_invoice_format.docx (默认模板)'
            })
    except Exception as e:
        return jsonify({'error': f'检查模板状态时出错: {str(e)}'}), 500

@app.route('/health')
def health_check():
    """健康检查端点"""
    return jsonify({'message': '服务正常运行', 'status': 'healthy'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
