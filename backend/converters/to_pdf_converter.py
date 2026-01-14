"""各种格式转PDF的转换器"""
import os
from typing import Optional
from .base_converter import BaseConverter


class WordToPDFConverter(BaseConverter):
    """Word转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将Word文档转换为PDF"""
        from docx2pdf import convert
        
        self.validate_file(input_file, ['.docx', '.doc'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        try:
            convert(input_file, output_file)
            return output_file
        except Exception as e:
            raise RuntimeError(f"Word转PDF失败: {str(e)}")


class PPTToPDFConverter(BaseConverter):
    """PowerPoint转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将PowerPoint转换为PDF"""
        try:
            from comtypes import client
        except ImportError:
            raise ImportError("需要安装comtypes库: pip install comtypes")
        
        self.validate_file(input_file, ['.pptx', '.ppt'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        try:
            # 使用PowerPoint COM接口
            powerpoint = client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            
            abs_input = os.path.abspath(input_file)
            abs_output = os.path.abspath(output_file)
            
            deck = powerpoint.Presentations.Open(abs_input, WithWindow=False)
            deck.SaveAs(abs_output, 32)  # 32 = PDF格式
            deck.Close()
            powerpoint.Quit()
            
            return output_file
        except Exception as e:
            raise RuntimeError(f"PPT转PDF失败: {str(e)}")


class ExcelToPDFConverter(BaseConverter):
    """Excel转PDF转换器"""
    
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """将Excel转换为PDF"""
        try:
            from comtypes import client
        except ImportError:
            raise ImportError("需要安装comtypes库: pip install comtypes")
        
        self.validate_file(input_file, ['.xlsx', '.xls'])
        output_file = self.get_output_path(input_file, '.pdf', output_file)
        
        try:
            excel = client.CreateObject("Excel.Application")
            excel.Visible = 0
            
            abs_input = os.path.abspath(input_file)
            abs_output = os.path.abspath(output_file)
            
            wb = excel.Workbooks.Open(abs_input)
            wb.ExportAsFixedFormat(0, abs_output)  # 0 = PDF格式
            wb.Close()
            excel.Quit()
            
            return output_file
        except Exception as e:
            raise RuntimeError(f"Excel转PDF失败: {str(e)}")


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
