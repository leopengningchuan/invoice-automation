import os
import pandas as pd
from datetime import datetime
import re
import zipfile
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
import warnings
import time
import pythoncom

# 导入自定义工具
from utils.docx_manipulate import populate_docx_table, convert_docx_pdf

warnings.filterwarnings('ignore')

app = Flask(__name__)
CORS(app)

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'
TEMP_FOLDER = 'temp'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# 确保文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_inv_info(inv_info, inv_no, sales_tax_rate=0.1):
    """获取发票信息的函数"""
    # 使用特定发票的信息
    use_info = inv_info[inv_info['Invoice No.'] == inv_no].reset_index(drop=True)
    item_dict = {}
    
    # 发票信息
    item_dict['CUSTOMER'] = use_info['Customer'].unique()[0]
    item_dict['CUSTOMER_ADDRESS1'] = use_info['Customer Address1'].unique()[0]
    item_dict['CUSTOMER_ADDRESS2'] = use_info['Customer Address2'].unique()[0]
    item_dict['INV_NO'] = use_info['Invoice No.'].unique()[0]
    item_dict['PAYMENT_TERMS'] = use_info['Payment Terms'].unique()[0]
    item_dict['DOC_DATE'] = str(pd.to_datetime(use_info['Invoice Date'].unique()[0]).date())
    item_dict['DUE_DATE'] = str(pd.to_datetime(use_info['Invoice Date'].unique()[0]).date() + pd.Timedelta(days=inv_info['Payment Terms'].unique()[0]))
    item_dict['SUB_AMOUNT'] = 0
    
    # 产品信息，最多15个项目
    for i in range(1, 16):
        try:
            # 对于非空项目，获取信息
            item_dict['ITEM' + str(i)] = use_info.loc[i-1, 'Item']
            item_dict['DETAIL' + str(i)] = use_info.loc[i-1, 'Detail']
            item_dict['UNITPRICE' + str(i)] = use_info.loc[i-1, 'Unit Price']
            item_dict['QUAN' + str(i)] = use_info.loc[i-1, 'Quantity']
            item_dict['AMT' + str(i)] = use_info.loc[i-1, 'Unit Price'] * use_info.loc[i-1, 'Quantity']
            
            # 获取小计总和
            item_dict['SUB_AMOUNT'] += item_dict['AMT' + str(i)]
            
        except:
            # 对于空项目，输入空信息
            item_dict['ITEM' + str(i)] = ""
            item_dict['DETAIL' + str(i)] = ""
            item_dict['UNITPRICE' + str(i)] = ""
            item_dict['QUAN' + str(i)] = ""
            item_dict['AMT' + str(i)] = ""
    
    # 获取税额和总金额
    item_dict['TAX_AMOUNT'] = round(item_dict['SUB_AMOUNT'] * sales_tax_rate, 2)
    item_dict['TOTAL_AMOUNT'] = item_dict['SUB_AMOUNT'] + item_dict['TAX_AMOUNT']
    
    # 将所有单价、金额转换为000,000.00格式
    for k in item_dict.keys():
        if any(x in k for x in ('UNITPRICE', 'AMT', 'AMOUNT', 'QUAN', 'PAYMENT_TERMS')):
            try:
                if any(x in k for x in ('UNITPRICE', 'AMT', 'AMOUNT')):
                    item_dict[k] = '{0:,.2f}'.format(item_dict[k])
                else: 
                    item_dict[k] = '{0:,}'.format(item_dict[k])
            except:
                pass
            
    return item_dict

def convert_docx_to_pdf_safe(docx_path, pdf_path):
    """安全的DOCX到PDF转换，包含COM初始化"""
    try:
        # 初始化COM组件
        pythoncom.CoInitialize()
        
        # 转换为PDF
        convert_docx_pdf(docx_path, keep=False)
        
        # 检查PDF是否生成
        if os.path.exists(pdf_path):
            return True
        else:
            # 如果PDF路径不对，查找生成的PDF
            pdf_dir = os.path.dirname(pdf_path)
            for file in os.listdir(pdf_dir):
                if file.endswith('.pdf') and os.path.splitext(os.path.basename(docx_path))[0] in file:
                    actual_pdf = os.path.join(pdf_dir, file)
                    if actual_pdf != pdf_path:
                        shutil.move(actual_pdf, pdf_path)
                    return True
        
        return False
        
    except Exception as e:
        print(f"PDF转换错误: {e}")
        return False
    finally:
        # 清理COM组件
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route('/')
def index():
    return render_template('index.html')

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
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        # 检查文件类型
        if not allowed_file(file.filename):
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
        filename = secure_filename(file.filename)
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
                
                # 替换Word模板中的发票信息
                populate_docx_table(item_dict, 'inv_template.docx', docx_path)
                
                # 使用安全的PDF转换方法
                if convert_docx_to_pdf_safe(docx_path, pdf_path):
                    # 删除DOCX文件
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
                    
                    # 添加生成的PDF到列表
                    if os.path.exists(pdf_path):
                        generated_files.append(pdf_filename)
                else:
                    return jsonify({'error': f'转换PDF失败: {item_dict["INV_NO"]}'}), 500
                
            except Exception as e:
                return jsonify({'error': f'生成发票 {inv_no} 时出错: {str(e)}'}), 500
        
        # 创建ZIP文件包含所有生成的PDF
        zip_filename = f"invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(app.config['TEMP_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in generated_files:
                pdf_path = os.path.join(temp_dir, pdf_file)
                zipf.write(pdf_path, pdf_file)
        
        # 清理临时文件夹
        shutil.rmtree(temp_dir)
        
        # 删除上传的Excel文件
        os.remove(filepath)
        
        # 计算总处理时间
        total_time = time.time() - start_time
        
        return jsonify({
            'message': f'发票生成成功 (税率: {sales_tax_rate*100:.1f}%)',
            'generated_files': generated_files,
            'zip_file': zip_filename,
            'download_url': f'/download/{zip_filename}',
            'processing_time': f'{total_time:.1f}秒',
            'tax_rate': f'{sales_tax_rate*100:.1f}%'
        })
        
    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载生成的ZIP文件"""
    try:
        file_path = os.path.join(app.config['TEMP_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': '文件不存在'}), 404
    except Exception as e:
        return jsonify({'error': f'下载文件时出错: {str(e)}'}), 500

@app.route('/health')
def health_check():
    """健康检查端点"""
    return jsonify({'status': 'healthy', 'message': 'Invoice自动化服务运行正常'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 