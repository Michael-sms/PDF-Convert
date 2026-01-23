"""Flask Web API服务"""
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import sys
import uuid
from datetime import datetime
import time

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.converters import (
    WordToPDFConverter, PPTToPDFConverter, ExcelToPDFConverter,
    ImageToPDFConverter, HTMLToPDFConverter,
    PDFToWordConverter, PDFToPPTConverter, PDFToImageConverter,
    PDFToExcelConverter
)

app = Flask(__name__, 
           template_folder='../frontend/templates',
           static_folder='../frontend/static')

# 配置
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB最大文件大小

# 使用绝对路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')

# 创建必要的文件夹
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 转换器映射
CONVERTERS = {
    'word2pdf': WordToPDFConverter(),
    'ppt2pdf': PPTToPDFConverter(),
    'excel2pdf': ExcelToPDFConverter(),
    'img2pdf': ImageToPDFConverter(),
    'html2pdf': HTMLToPDFConverter(),
    'pdf2word': PDFToWordConverter(),
    'pdf2ppt': PDFToPPTConverter(),
    'pdf2img': PDFToImageConverter(),
    'pdf2excel': PDFToExcelConverter(),
}

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {
    'word2pdf': {'.docx', '.doc'},
    'ppt2pdf': {'.pptx', '.ppt'},
    'excel2pdf': {'.xlsx', '.xls'},
    'img2pdf': {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'},
    'html2pdf': {'.html', '.htm'},
    'pdf2word': {'.pdf'},
    'pdf2ppt': {'.pdf'},
    'pdf2img': {'.pdf'},
    'pdf2excel': {'.pdf'},
}


def allowed_file(filename, conversion_type):
    """检查文件是否允许"""
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS.get(conversion_type, set())


def cleanup_old_files(folder_path, max_age_hours=24):
    """清理超过指定时间的文件"""
    try:
        if not os.path.exists(folder_path):
            return
        
        current_time = time.time()
        max_age_seconds = max_age_hours * 3600
        
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                file_age = current_time - os.path.getmtime(file_path)
                if file_age > max_age_seconds:
                    try:
                        os.remove(file_path)
                        print(f"已清理旧文件: {filename}")
                    except Exception as e:
                        print(f"清理文件失败 {filename}: {e}")
    except Exception as e:
        print(f"清理文件夹失败: {e}")


@app.route('/')
def index():
    """首页"""
    # 每次访问首页时清理超过24小时的旧文件
    cleanup_old_files(app.config['UPLOAD_FOLDER'], max_age_hours=1)
    cleanup_old_files(app.config['OUTPUT_FOLDER'], max_age_hours=24)
    return render_template('index.html')


@app.route('/api/convert', methods=['POST'])
def convert_file():
    """文件转换API"""
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'error': '没有文件上传'}), 400
        
        file = request.files['file']
        conversion_type = request.form.get('type')
        
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        if not conversion_type or conversion_type not in CONVERTERS:
            return jsonify({'error': '无效的转换类型'}), 400
        
        if not allowed_file(file.filename, conversion_type):
            return jsonify({'error': f'不支持的文件格式'}), 400
        
        # 保存上传的文件
        filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4())[:8]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        input_filename = f"{timestamp}_{unique_id}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        
        print(f"保存上传文件到: {input_path}")
        file.save(input_path)
        
        # 执行转换
        try:
            converter = CONVERTERS[conversion_type]
            
            # 生成输出文件路径
            output_ext_map = {
                'word2pdf': '.pdf', 'ppt2pdf': '.pdf', 'excel2pdf': '.pdf',
                'img2pdf': '.pdf', 'html2pdf': '.pdf',
                'pdf2word': '.docx', 'pdf2ppt': '.pptx',
                'pdf2img': '.jpg', 'pdf2excel': '.xlsx',
            }
            output_ext = output_ext_map[conversion_type]
            output_filename = f"{os.path.splitext(input_filename)[0]}_converted{output_ext}"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            
            print(f"输出文件将保存到: {output_path}")
            
            # 转换文件
            result = converter.convert(input_path, output_path)
            
            print(f"转换完成，实际输出: {result}")
            print(f"文件是否存在: {os.path.exists(output_path)}")
            
            # 返回下载链接
            return jsonify({
                'success': True,
                'message': '转换成功',
                'download_url': f'/api/download/{output_filename}',
                'filename': output_filename
            })
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            return jsonify({'error': f'转换失败: {str(e)}'}), 500
        
        finally:
            # 清理上传的文件
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except:
                    pass
    
    except Exception as e:
        return jsonify({'error': f'处理请求失败: {str(e)}'}), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    """下载文件"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        # 调试信息
        print(f"尝试下载文件: {filename}")
        print(f"完整路径: {file_path}")
        print(f"文件是否存在: {os.path.exists(file_path)}")
        
        if not os.path.exists(file_path):
            # 列出outputs目录下的所有文件
            if os.path.exists(app.config['OUTPUT_FOLDER']):
                files = os.listdir(app.config['OUTPUT_FOLDER'])
                print(f"outputs目录下的文件: {files}")
            return jsonify({'error': '文件不存在', 'path': file_path}), 404
        
        # 下载文件后，设置回调来删除文件
        response = send_file(file_path, as_attachment=True, download_name=filename)
        
        # 可选：延迟删除文件（下载完成后）
        # 这里暂时不删除，让文件保留一段时间以便用户多次下载
        
        return response
    
    except Exception as e:
        print(f"下载错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'下载失败: {str(e)}'}), 500


@app.route('/api/info')
def get_info():
    """获取支持的转换类型信息"""
    info = {
        'to_pdf': {
            'word2pdf': {'name': 'Word转PDF', 'formats': ['.docx', '.doc']},
            'ppt2pdf': {'name': 'PPT转PDF', 'formats': ['.pptx', '.ppt']},
            'excel2pdf': {'name': 'Excel转PDF', 'formats': ['.xlsx', '.xls']},
            'img2pdf': {'name': '图片转PDF', 'formats': ['.jpg', '.jpeg', '.png', '.bmp', '.gif']},
            'html2pdf': {'name': 'HTML转PDF', 'formats': ['.html', '.htm']},
        },
        'from_pdf': {
            'pdf2word': {'name': 'PDF转Word', 'formats': ['.pdf']},
            'pdf2ppt': {'name': 'PDF转PPT', 'formats': ['.pdf']},
            'pdf2img': {'name': 'PDF转图片', 'formats': ['.pdf']},
            'pdf2excel': {'name': 'PDF转Excel', 'formats': ['.pdf']},
        }
    }
    return jsonify(info)


if __name__ == '__main__':
    print("=" * 50)
    print("PDF转换器 Web服务启动中...")
    print("访问地址: http://localhost:5000")
    print("-" * 50)
    print(f"上传文件夹: {app.config['UPLOAD_FOLDER']}")
    print(f"输出文件夹: {app.config['OUTPUT_FOLDER']}")
    print(f"上传文件夹存在: {os.path.exists(app.config['UPLOAD_FOLDER'])}")
    print(f"输出文件夹存在: {os.path.exists(app.config['OUTPUT_FOLDER'])}")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=5000)
