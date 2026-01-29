# PDF转换器

一个功能强大的PDF转换工具，支持多种文档格式互转，提供命令行界面和现代化的Web界面。

## 📸 界面展示

### 主界面
![界面展示](docx/images/界面展示.png)

**赛博科技风格 (Cyberpunk Tech Style)** 的Web界面：
- 🌌 深空网格背景 + 全息扫描线效果
- ⚡ 电光青霓虹主色调
- 🪟 暗黑玻璃拟态 (Dark Glass Morphism)
- ✨ 浮动光斑动画 + 按钮光泽特效
- 📱 支持拖拽上传，响应式设计

### 转换成功
![转换成功展示](docx/images/转换成功展示.png)

转换完成后可直接下载文件，界面提供清晰的状态反馈。

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
git clone https://github.com/Michael-sms/PDF-Convert.git
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
- **现代CSS3** - 赛博科技风格、渐变、动画、Glass Morphism
- **HTML5** - 拖拽上传、文件API
- **设计风格** - Cyberpunk/Sci-Fi Tech 深色主题

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

### Q: Word/PPT/Excel转PDF时遇到500错误或"Failed to fetch"？
**A:** v2.0.2已采用子进程隔离技术大幅改善此问题。请确保：
1. 已安装 **Microsoft Office**（不支持WPS）
2. 首次运行前**手动打开一次**Word/PowerPoint/Excel，关闭所有欢迎弹窗和激活提示
3. 重启Flask服务：`python backend/app.py`
4. 查看终端输出的 `[Excel转PDF]` 等日志获取详细错误信息
5. 如果问题持续，运行 `python check_word_convert.py` 进行诊断

详细排查步骤请参考 [故障排除指南](TROUBLESHOOTING.md)。

### Q: 转换超时（2分钟）？
**A:** 通常是Office应用弹出了对话框（如激活提示、宏安全警告等）。请手动打开对应的Office程序，确认没有任何弹窗后关闭，再重试转换。

### Q: HTML转PDF失败？
**A:** 需要安装wkhtmltopdf并添加到PATH环境变量。

### Q: PDF转图片失败？
**A:** 需要安装Poppler for Windows并添加到PATH环境变量。

### Q: PDF转Excel失败？
**A:** 需要安装Java运行环境，并确保表格在PDF中格式规范。

### Q: Web界面转换超时？
**A:** 大文件建议使用命令行界面，或增加Flask超时时间。

## 📝 更新日志

### v2.0.2 (2026-01-29)
- 🎨 **全新赛博科技风格UI** - 深空网格背景、电光青霓虹配色、暗黑玻璃拟态
- ⚡ 添加全息扫描线、浮动光斑动画、按钮霓虹辉光效果
- 🛡️ **子进程隔离技术** - Word/PPT/Excel转PDF使用独立子进程执行
- 🐛 修复COM组件崩溃导致Flask主进程终止的问题
- ⏱️ 添加转换超时保护（2分钟），防止弹窗阻塞
- 📊 增强错误日志输出，显示详细的转换步骤信息
- 🔧 优化Excel转PDF的COM调用，添加DisplayAlerts=False

### v2.0.1 (2026-01-23)
- 🐛 修复Flask多线程环境下Word/PPT转PDF的COM组件初始化错误
- 🔧 添加 `pythoncom.CoInitialize()` 和 `CoUninitialize()` 确保COM线程安全
- 🎨 优化Web界面为现代科技风格，采用动态渐变背景
- 🎨 实现玻璃拟态(Glass Morphism)设计风格
- ✨ 添加浮动动画效果和按钮光泽特效
- 📚 新增故障排除指南 `TROUBLESHOOTING.md`
- 🛠️ 添加诊断工具 `check_word_convert.py`
- 📖 改进错误日志输出，便于调试

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
[Michael-sms](https://github.com/Michael-sms),有任何建议可题issue或用邮箱联系我。
**享受便捷的文档转换体验！** 🎉
