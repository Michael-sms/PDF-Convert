# PDF转换器故障排除指南

如果您遇到 500 错误或转换失败，请按照以下步骤检查：

## 1. 检查 Microsoft Office 安装
**重要**: Word 转 PDF 和 PPT 转 PDF 功能依赖于 Microsoft Office (Word 和 PowerPoint)。
- **WPS Office**: 不支持。必须是 Microsoft Office。
- **Web版/云端Office**: 不支持。必须是本地安装版。
- **已激活**: 确保您的 Office 已激活且能正常打开，没有弹窗干扰。

## 2. 第一次运行前的检查
当程序尝试调用 Word 时，如果 Word 弹出了"欢迎使用"或"激活提示"窗口，程序会卡住或报错。
- 请手动打开一次 Word 和 PowerPoint。
- 确保关闭所有弹窗。
- 确保能够正常新建和保存文档。

## 3. 使用诊断工具
如果您不确定问题出在哪里，请运行项目根目录下的诊断脚本：
```powershell
python check_word_convert.py
```
该脚本会：
1. 检查必要的依赖库
2. 创建一个测试 Word 文档
3. 尝试将其转换为 PDF
4. 输出详细的错误信息

## 4. 常见错误代码
- **500 Internal Server Error**: 通常是因为 Flask 多线程模式下 COM 组件初始化失败。这是一个已知问题，我们已经在最新代码中修复了它 (添加了 `pythoncom.CoInitialize()`)。
- **ImportError**: 缺少依赖库，请运行 `pip install -r requirements.txt`。
- **PermissionError**: 文件被占用。请确保转换时不要用 Word 打开该文件。

## 5. 重启服务
修改代码后，请务必重启 Flask 服务：
1. 在终端按 `Ctrl+C` 停止服务
2. 再次运行 `python backend/app.py`
