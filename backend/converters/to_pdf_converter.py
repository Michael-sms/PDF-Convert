"""各种格式转PDF的转换器"""
import os
from typing import Optional
from .base_converter import BaseConverter


class WordToPDFConverter(BaseConverter):
    """Word转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将Word文档转换为PDF"""
        import subprocess
        import sys
        
        self.validate_file(input_file, ['.docx', '.doc'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        abs_input = os.path.abspath(input_file)
        abs_output = os.path.abspath(output_file)
        
        # 使用子进程执行，防止COM崩溃影响主进程
        script = f'''
import pythoncom
import sys
pythoncom.CoInitialize()
try:
    from docx2pdf import convert
    convert(r"{abs_input}", r"{abs_output}")
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {{e}}", file=sys.stderr)
    sys.exit(1)
finally:
    pythoncom.CoUninitialize()
'''
        
        try:
            result = subprocess.run(
                [sys.executable, '-c', script],
                capture_output=True,
                text=True,
                timeout=120  # 2分钟超时
            )
            
            if result.returncode != 0 or "SUCCESS" not in result.stdout:
                error_msg = result.stderr.strip() if result.stderr else "未知错误"
                raise RuntimeError(f"Word转PDF失败: {error_msg}")
            
            return output_file
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("Word转PDF超时，可能是Word弹出了对话框。请手动打开Word确认无弹窗后重试。")
        except Exception as e:
            raise RuntimeError(f"Word转PDF失败: {str(e)}")


class PPTToPDFConverter(BaseConverter):
    """PowerPoint转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将PowerPoint转换为PDF"""
        import subprocess
        import sys
        
        self.validate_file(input_file, ['.pptx', '.ppt'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        abs_input = os.path.abspath(input_file)
        abs_output = os.path.abspath(output_file)
        
        # 使用子进程执行，防止COM崩溃影响主进程
        script = f'''
import pythoncom
import sys
pythoncom.CoInitialize()
try:
    from comtypes import client
    powerpoint = client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    
    deck = powerpoint.Presentations.Open(r"{abs_input}", WithWindow=False)
    deck.SaveAs(r"{abs_output}", 32)
    deck.Close()
    powerpoint.Quit()
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {{e}}", file=sys.stderr)
    sys.exit(1)
finally:
    pythoncom.CoUninitialize()
'''
        
        try:
            result = subprocess.run(
                [sys.executable, '-c', script],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode != 0 or "SUCCESS" not in result.stdout:
                error_msg = result.stderr.strip() if result.stderr else "未知错误"
                raise RuntimeError(f"PPT转PDF失败: {error_msg}")
            
            return output_file
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("PPT转PDF超时，可能是PowerPoint弹出了对话框。请手动打开PowerPoint确认无弹窗后重试。")
        except Exception as e:
            raise RuntimeError(f"PPT转PDF失败: {str(e)}")


class ExcelToPDFConverter(BaseConverter):
    """Excel转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将Excel转换为PDF"""
        import subprocess
        import sys
        
        self.validate_file(input_file, ['.xlsx', '.xls'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        abs_input = os.path.abspath(input_file)
        abs_output = os.path.abspath(output_file)
        
        # 使用子进程来执行转换，避免COM问题导致主进程崩溃
        script = f'''
import pythoncom
import sys
import traceback

pythoncom.CoInitialize()
excel = None
wb = None
try:
    from comtypes import client
    print("正在启动Excel...", file=sys.stderr)
    excel = client.CreateObject("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    print("正在打开文件...", file=sys.stderr)
    wb = excel.Workbooks.Open(r"{abs_input}", ReadOnly=True)
    
    print("正在导出PDF...", file=sys.stderr)
    wb.ExportAsFixedFormat(0, r"{abs_output}")
    
    print("SUCCESS")
except Exception as e:
    print(f"FAILED: {{type(e).__name__}}: {{e}}", file=sys.stderr)
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)
finally:
    try:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
    except:
        pass
    pythoncom.CoUninitialize()
'''
        
        try:
            print(f"[Excel转PDF] 开始转换: {abs_input}")
            result = subprocess.run(
                [sys.executable, '-c', script],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            print(f"[Excel转PDF] 返回码: {result.returncode}")
            if result.stdout:
                print(f"[Excel转PDF] stdout: {result.stdout}")
            if result.stderr:
                print(f"[Excel转PDF] stderr: {result.stderr}")
            
            if result.returncode != 0 or "SUCCESS" not in result.stdout:
                error_msg = result.stderr.strip() if result.stderr else "子进程执行失败"
                raise RuntimeError(error_msg)
            
            if not os.path.exists(abs_output):
                raise RuntimeError("转换完成但输出文件不存在")
                
            return output_file
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("Excel转PDF超时(2分钟)，请检查Excel是否弹出了对话框")
        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(f"执行出错: {str(e)}")


class ImageToPDFConverter(BaseConverter):
    """图片转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将图片转换为PDF"""
        try:
            from PIL import Image
        except ImportError:
            raise ImportError("需要安装Pillow库: pip install Pillow")
        
        self.validate_file(input_file, ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        try:
            image = Image.open(input_file)
            
            # 转换RGBA到RGB
            if image.mode == 'RGBA':
                background = Image.new('RGB', image.size, (255, 255, 255))
                background.paste(image, mask=image.split()[3])
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
            
            image.save(output_file, 'PDF', resolution=100.0)
            return output_file
        except Exception as e:
            raise RuntimeError(f"图片转PDF失败: {str(e)}")


class HTMLToPDFConverter(BaseConverter):
    """HTML转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将HTML转换为PDF"""
        try:
            import pdfkit
        except ImportError:
            raise ImportError("需要安装pdfkit库: pip install pdfkit")
        
        self.validate_file(input_file, ['.html', '.htm'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        try:
            pdfkit.from_file(input_file, output_file)
            return output_file
        except Exception as e:
            raise RuntimeError(f"HTML转PDF失败: {str(e)}\n提示: 需要安装wkhtmltopdf (https://wkhtmltopdf.org/)")
