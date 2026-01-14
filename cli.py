"""命令行界面"""
import argparse
import os
import sys

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.converters import (
    WordToPDFConverter, PPTToPDFConverter, ExcelToPDFConverter,
    ImageToPDFConverter, HTMLToPDFConverter,
    PDFToWordConverter, PDFToPPTConverter, PDFToImageConverter,
    PDFToExcelConverter
)


class PDFConverterCLI:
    """PDF转换器命令行界面"""
    
    def __init__(self):
        self.converters = {
            # 转为PDF
            'word2pdf': WordToPDFConverter(),
            'ppt2pdf': PPTToPDFConverter(),
            'excel2pdf': ExcelToPDFConverter(),
            'img2pdf': ImageToPDFConverter(),
            'html2pdf': HTMLToPDFConverter(),
            # PDF转其他
            'pdf2word': PDFToWordConverter(),
            'pdf2ppt': PDFToPPTConverter(),
            'pdf2img': PDFToImageConverter(),
            'pdf2excel': PDFToExcelConverter(),
        }
    
    def convert_file(self, conversion_type: str, input_file: str, output_file: str = None):
        """转换单个文件"""
        if conversion_type not in self.converters:
            print(f"错误: 不支持的转换类型 '{conversion_type}'")
            print(f"支持的转换类型: {', '.join(self.converters.keys())}")
            return False
        
        try:
            converter = self.converters[conversion_type]
            result = converter.convert(input_file, output_file)
            print(f"✓ 转换成功: {result}")
            return True
        except Exception as e:
            print(f"✗ 转换失败: {str(e)}")
            return False
    
    def batch_convert(self, conversion_type: str, input_dir: str):
        """批量转换文件"""
        if not os.path.isdir(input_dir):
            print(f"错误: '{input_dir}' 不是有效的目录")
            return
        
        # 根据转换类型确定要处理的文件扩展名
        extension_map = {
            'word2pdf': ['.docx', '.doc'],
            'ppt2pdf': ['.pptx', '.ppt'],
            'excel2pdf': ['.xlsx', '.xls'],
            'img2pdf': ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'],
            'html2pdf': ['.html', '.htm'],
            'pdf2word': ['.pdf'],
            'pdf2ppt': ['.pdf'],
            'pdf2img': ['.pdf'],
            'pdf2excel': ['.pdf'],
        }
        
        extensions = extension_map.get(conversion_type, [])
        files_converted = 0
        files_failed = 0
        
        print(f"开始批量转换 '{input_dir}' 中的文件...")
        
        for filename in os.listdir(input_dir):
            file_path = os.path.join(input_dir, filename)
            
            if not os.path.isfile(file_path):
                continue
            
            ext = os.path.splitext(filename)[1].lower()
            if ext in extensions:
                print(f"\n处理: {filename}")
                if self.convert_file(conversion_type, file_path):
                    files_converted += 1
                else:
                    files_failed += 1
        
        print(f"\n批量转换完成!")
        print(f"成功: {files_converted} 个文件")
        print(f"失败: {files_failed} 个文件")
    
    def run(self):
        """运行命令行界面"""
        parser = argparse.ArgumentParser(
            description='PDF转换器 - 支持多种格式互转',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
转换类型:
  转为PDF:
    word2pdf    - Word转PDF
    ppt2pdf     - PowerPoint转PDF
    excel2pdf   - Excel转PDF
    img2pdf     - 图片转PDF
    html2pdf    - HTML转PDF
  
  PDF转其他:
    pdf2word    - PDF转Word
    pdf2ppt     - PDF转PowerPoint
    pdf2img     - PDF转图片
    pdf2excel   - PDF转Excel

示例:
  # 单文件转换
  python cli.py word2pdf document.docx
  python cli.py pdf2img report.pdf -o output.jpg
  
  # 批量转换
  python cli.py word2pdf -d ./documents
  python cli.py pdf2word -d ./pdfs
            """
        )
        
        parser.add_argument('type', 
                          choices=list(self.converters.keys()),
                          help='转换类型')
        parser.add_argument('input', 
                          nargs='?',
                          help='输入文件路径')
        parser.add_argument('-o', '--output', 
                          help='输出文件路径（可选）')
        parser.add_argument('-d', '--directory',
                          help='批量转换：输入文件夹路径')
        
        args = parser.parse_args()
        
        # 批量转换模式
        if args.directory:
            self.batch_convert(args.type, args.directory)
        # 单文件转换模式
        elif args.input:
            self.convert_file(args.type, args.input, args.output)
        else:
            parser.print_help()
            print("\n错误: 请指定输入文件或使用 -d 指定输入文件夹")


if __name__ == '__main__':
    cli = PDFConverterCLI()
    cli.run()
