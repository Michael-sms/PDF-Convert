# PDF转换器

一个功能强大的PDF转换工具，支持多种文档格式互转，提供命令行界面和现代化的Web界面。

## ✨ 功能特性

### 📄 转为PDF
- **Word转PDF** - 支持 .docx, .doc 格式
- **PPT转PDF** - 支持 .pptx, .ppt 格式
- **Excel转PDF** - 支持 .xlsx, .xls 格式
- **图片转PDF** - 支持 .jpg, .png, .bmp, .gif, .tiff 格式
- **HTML转PDF** - 支持 .html, .htm 格式

### 📑 PDF转其他格式
- **PDF转Word** - 转换为 .docx 格式
- **PDF转PPT** - 转换为 .pptx 格式
- **PDF转图片** - 转换为 .jpg 格式
- **PDF转Excel** - 转换为 .xlsx 格式

## 🚀 快速开始

### 环境要求

- Python 3.8 或更高版本
- Windows操作系统（部分功能需要）

### 安装步骤

1. **克隆或下载项目**
```bash
git clone <repository-url>
cd Convert2PDF
```

2. **安装Python依赖**
```bash
pip install -r requirements.txt
```

3. **安装外部依赖（根据需要）**

#### Windows用户 - Office转换功能
- 安装 Microsoft Office（Word、PowerPoint、Excel）
- 用于 PPT转PDF 和 Excel转PDF 功能

#### HTML转PDF功能
- 下载安装 [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html)
- 将 wkhtmltopdf 添加到系统PATH环境变量

#### PDF转图片/PDF转PPT功能
- 下载 [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/)
- 解压并将 `bin` 目录添加到系统PATH环境变量

#### PDF转Excel功能
- 安装 [Java运行环境](https://www.oracle.com/java/technologies/downloads/)

## 💻 使用方法

### 方式一：Web界面（推荐）

1. **启动Web服务**
```bash
python backend/app.py
```

2. **打开浏览器访问**
```
http://localhost:5000
```

3. **使用步骤**
   - 选择转换类型（转为PDF 或 PDF转其他）
   - 点击具体的转换选项（如 Word→PDF）
   - 上传文件（支持拖拽）
   - 点击"开始转换"
   - 转换完成后下载文件

### 方式二：命令行界面

#### 单文件转换

```bash
# Word转PDF
python cli.py word2pdf document.docx

# PDF转Word
python cli.py pdf2word report.pdf

# 图片转PDF
python cli.py img2pdf photo.jpg

# 指定输出文件名
python cli.py word2pdf input.docx -o output.pdf
```

#### 批量转换

```bash
# 批量转换文件夹中的所有Word文档为PDF
python cli.py word2pdf -d ./documents

# 批量转换文件夹中的所有PDF为Word
python cli.py pdf2word -d ./pdfs
```

#### 查看帮助

```bash
python cli.py --help
```

## 📁 项目结构

```
Convert2PDF/
├── backend/                    # 后端代码
│   ├── converters/            # 转换器模块
│   │   ├── base_converter.py  # 基础转换器类
│   │   ├── to_pdf_converter.py    # 转为PDF的转换器
│   │   ├── from_pdf_converter.py  # PDF转其他格式的转换器
│   │   └── __init__.py
│   └── app.py                 # Flask Web服务
├── frontend/                   # 前端代码
│   ├── templates/             # HTML模板
│   │   └── index.html         # 主页面
│   └── static/                # 静态资源
├── cli.py                     # 命令行界面
├── requirements.txt           # Python依赖
├── README.md                  # 项目文档
└── word2pdf.py               # 旧版本（已废弃）
```

## 🛠️ 技术栈

### 后端
- **Flask** - Web框架
- **docx2pdf** - Word转PDF
- **pdf2docx** - PDF转Word
- **Pillow** - 图片处理
- **pdfkit** - HTML转PDF
- **pdf2image** - PDF转图片
- **python-pptx** - PPT处理
- **tabula-py** - PDF表格提取
- **comtypes** - Windows Office自动化

### 前端
- **原生JavaScript** - 交互逻辑
- **现代CSS** - 渐变、动画、响应式设计
- **HTML5** - 拖拽上传、文件API

## ⚠️ 注意事项

1. **文件大小限制**
   - Web界面最大支持50MB文件
   - 命令行无大小限制

2. **转换质量**
   - Word/PPT/Excel转PDF需要安装Microsoft Office
   - PDF转Word保持格式可能有偏差
   - PDF转PPT会将每页转为图片

3. **性能**
   - 大文件转换可能需要较长时间
   - 批量转换建议使用命令行界面

4. **外部依赖**
   - 部分功能需要安装额外软件（见安装步骤）
   - 确保软件已添加到PATH环境变量

## 🔧 常见问题

### Q: PPT转PDF或Excel转PDF失败？
**A:** 确保已安装Microsoft Office，并且comtypes库已正确安装。

### Q: HTML转PDF失败？
**A:** 需要安装wkhtmltopdf并添加到PATH环境变量。

### Q: PDF转图片失败？
**A:** 需要安装Poppler for Windows并添加到PATH环境变量。

### Q: PDF转Excel失败？
**A:** 需要安装Java运行环境，并确保表格在PDF中格式规范。

### Q: Web界面转换超时？
**A:** 大文件建议使用命令行界面，或增加Flask超时时间。

## 📝 更新日志

### v2.0.0 (2026-01-14)
- ✨ 全面重构代码架构
- ✨ 新增9种格式转换支持
- ✨ 添加现代化Web界面
- ✨ 支持拖拽上传
- ✨ 添加批量转换功能
- 🔧 改进错误处理
- 📚 完善文档说明

### v1.0.0
- 基础Word转PDF功能
- Tkinter图形界面

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 👨‍💻 作者

PDF转换器项目团队

---

**享受便捷的文档转换体验！** 🎉
