# PDF转换器 - 快速入门指南

欢迎使用PDF转换器！本指南将帮助您快速上手。

## 🎯 第一次使用

### 1. 安装依赖

打开命令行（PowerShell或CMD），在项目目录下执行：

```bash
pip install -r requirements.txt
```

这将安装所有必需的Python库。

### 2. 选择使用方式

您有两种使用方式可选：

#### 方式A：Web界面（推荐新手）

**双击运行：**
- 找到项目中的 `start_web.ps1` 文件
- 右键 → "使用PowerShell运行"
- 浏览器会自动打开 http://localhost:5000

**或者手动启动：**
```bash
python backend/app.py
```

#### 方式B：命令行界面（适合批量处理）

```bash
# 查看帮助
python cli.py --help

# 转换单个文件
python cli.py word2pdf 文档.docx

# 批量转换文件夹
python cli.py word2pdf -d ./文档文件夹
```

## 📋 支持的转换类型

### 转为PDF
- `word2pdf` - Word转PDF
- `ppt2pdf` - PPT转PDF
- `excel2pdf` - Excel转PDF
- `img2pdf` - 图片转PDF
- `html2pdf` - HTML转PDF

### PDF转其他
- `pdf2word` - PDF转Word
- `pdf2ppt` - PDF转PPT
- `pdf2img` - PDF转图片
- `pdf2excel` - PDF转Excel

## 💡 使用示例

### Web界面使用流程

1. 打开网页后，选择一个标签页：
   - "转为PDF" 或 "PDF转其他"

2. 点击具体的转换按钮：
   - 例如："Word → PDF"

3. 上传文件：
   - 点击上传区域选择文件
   - 或直接拖拽文件到上传区域

4. 点击"开始转换"按钮

5. 等待转换完成后，点击"下载文件"

### 命令行使用示例

```bash
# Word转PDF
python cli.py word2pdf 报告.docx
python cli.py word2pdf 报告.docx -o 输出.pdf

# PDF转Word
python cli.py pdf2word 文档.pdf

# 图片转PDF
python cli.py img2pdf 照片.jpg

# 批量转换
python cli.py word2pdf -d C:\Documents
python cli.py pdf2word -d C:\PDFs
```

## ⚙️ 额外软件安装（按需）

某些转换功能需要额外软件支持：

### PPT转PDF / Excel转PDF
- 需要：Microsoft Office
- 下载：https://www.microsoft.com/microsoft-365

### HTML转PDF
- 需要：wkhtmltopdf
- 下载：https://wkhtmltopdf.org/downloads.html
- 安装后添加到PATH环境变量

### PDF转图片 / PDF转PPT
- 需要：Poppler
- 下载：https://github.com/oschwartz10612/poppler-windows/releases/
- 解压后将bin目录添加到PATH环境变量

### PDF转Excel
- 需要：Java
- 下载：https://www.oracle.com/java/technologies/downloads/

## 🔍 常见问题快速解决

### Q1: 启动Web服务时报错
```
ModuleNotFoundError: No module named 'flask'
```
**解决：** 运行 `pip install -r requirements.txt`

### Q2: 转换失败提示缺少软件
**解决：** 查看上方"额外软件安装"章节，安装对应软件

### Q3: Web界面打不开
**解决：**
1. 检查是否正确启动了服务
2. 确认访问地址：http://localhost:5000
3. 尝试关闭防火墙

### Q4: 文件上传后没反应
**解决：**
1. 检查文件格式是否正确
2. 文件大小不要超过50MB
3. 查看浏览器控制台是否有错误

## 📞 获取帮助

遇到问题？

1. 查看完整文档：README.md
2. 检查错误提示信息
3. 提交Issue到GitHub仓库

## 🎉 开始使用吧！

现在您已经准备好了，开始愉快地转换文档吧！

**推荐第一次尝试：**
1. 运行 `python backend/app.py`
2. 打开浏览器访问 http://localhost:5000
3. 选择 "Word → PDF"
4. 上传一个.docx文件
5. 点击转换并下载

祝使用愉快！ 🚀
