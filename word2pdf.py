import subprocess

from docx2pdf import convert
import argparse
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class WordToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF工具")
        self.root.geometry("500x300")

        # 设置样式
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('TLabel', font=('Arial', 10))

        # 创建界面元素
        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.root, text="Word转PDF转换器", font=('Arial', 14, 'bold'))
        title_label.pack(pady=10)

        # 单文件转换区域
        file_frame = ttk.LabelFrame(self.root, text="单个文件转换", padding=10)
        file_frame.pack(pady=5, padx=10, fill=tk.X)

        self.file_path = tk.StringVar()
        ttk.Label(file_frame, text="选择Word文件:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.file_path, width=40).grid(row=1, column=0, padx=5)
        ttk.Button(file_frame, text="浏览...", command=self.browse_file).grid(row=1, column=1)
        ttk.Button(file_frame, text="转换", command=self.convert_single_file).grid(row=2, column=0, columnspan=2,
                                                                                   pady=5)

        # 批量转换区域
        dir_frame = ttk.LabelFrame(self.root, text="批量转换", padding=10)
        dir_frame.pack(pady=5, padx=10, fill=tk.X)

        self.dir_path = tk.StringVar()
        ttk.Label(dir_frame, text="选择文件夹:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(dir_frame, textvariable=self.dir_path, width=40).grid(row=1, column=0, padx=5)
        ttk.Button(dir_frame, text="浏览...", command=self.browse_dir).grid(row=1, column=1)
        ttk.Button(dir_frame, text="批量转换", command=self.convert_directory).grid(row=2, column=0, columnspan=2,
                                                                                    pady=5)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.BOTTOM, fill=tk.X)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)

    def browse_dir(self):
        dir_path = filedialog.askdirectory(title="选择包含Word文件的文件夹")
        if dir_path:
            self.dir_path.set(dir_path)

    def convert_single_file(self):
        input_file = self.file_path.get()
        if not input_file:
            messagebox.showwarning("警告", "请先选择Word文件")
            return

        self.status_var.set("正在转换单个文件...")
        self.root.update()

        result = self.convert_file(input_file)
        if result:
            messagebox.showinfo("成功", f"文件已转换为: {result}")
        self.status_var.set("就绪")

    def convert_directory(self):
        input_dir = self.dir_path.get()
        if not input_dir:
            messagebox.showwarning("警告", "请先选择文件夹")
            return

        self.status_var.set("正在批量转换文件...")
        self.root.update()

        try:
            count = len(self.convert_dir(input_dir))
            messagebox.showinfo("成功", f"批量转换完成，共转换了{count}个文件")
        except Exception as e:
            messagebox.showerror("错误", f"批量转换失败: {str(e)}")
        self.status_var.set("就绪")

    def convert_file(self, input_file):
        """转换单个文件为PDF"""
        if not input_file.lower().endswith('.docx'):
            messagebox.showerror("错误", "文件格式不正确，请选择.docx文件")
            return None

        file_name = os.path.splitext(input_file)[0]
        output_file = file_name + ".pdf"

        try:
            convert(input_file, output_file)
            return output_file
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            return None

    def convert_dir(self, input_dir):
        """转换目录下所有文件"""
        if not os.path.isdir(input_dir):
            messagebox.showerror("错误", "输入路径不存在或不是目录")
            return []

        output_files = []
        for file_name in os.listdir(input_dir):
            if file_name.lower().endswith('.docx'):
                input_file = os.path.join(input_dir, file_name)
                output_file = self.convert_file(input_file)
                if output_file:
                    output_files.append(output_file)

        return output_files
    '''
    def convert_to_word(self, input_file):
        """转换单个文件为Word"""
        if not input_file.lower().endswith('.pdf'):
            messagebox.showerror("错误", "文件格式不正确，请选择.pdf文件")
            return None

        file_name = os.path.splitext(input_file)[0]
        output_file = file_name + ".docx"

        try:
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'docx', input_file, '--outdir', os.path.dirname(output_file)], check=True)
            return output_file
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            return None
    '''


def command_line_interface():
    """命令行界面"""
    parser = argparse.ArgumentParser(description='Convert Word document to PDF.')
    parser.add_argument('input_file', type=str, nargs='?', help='docx文件路径或文件夹路径')
    args = parser.parse_args()

    if args.input_file:
        converter = WordToPDFConverter(tk.Tk())
        input_path = args.input_file
        if os.path.isfile(input_path):
            converter.convert_file(input_path)
        elif os.path.isdir(input_path):
            converter.convert_dir(input_path)
        else:
            print(f"输入路径{input_path}不存在或不是文件或目录")
    else:
        root = tk.Tk()
        app = WordToPDFConverter(root)
        root.mainloop()


if __name__ == '__main__':
    command_line_interface()



