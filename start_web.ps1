# 启动Web界面
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "PDF转换器 - Web服务" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""
Write-Host "正在启动Web服务..." -ForegroundColor Yellow
Write-Host "启动后请在浏览器访问: http://localhost:5000" -ForegroundColor Yellow
Write-Host ""
Write-Host "按 Ctrl+C 停止服务" -ForegroundColor Red
Write-Host ""

# 检查Python是否安装
try {
    $pythonVersion = python --version 2>&1
    Write-Host "检测到: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "错误: 未找到Python，请先安装Python 3.8+!" -ForegroundColor Red
    pause
    exit
}

# 检查依赖
Write-Host "检查依赖..." -ForegroundColor Yellow
$requirements = Get-Content "requirements.txt"
$missingPackages = @()

foreach ($req in $requirements) {
    if ($req -match '^([a-zA-Z0-9_-]+)==') {
        $package = $matches[1]
        $result = python -c "import $package" 2>&1
        if ($LASTEXITCODE -ne 0) {
            $missingPackages += $package
        }
    }
}

if ($missingPackages.Count -gt 0) {
    Write-Host "警告: 缺少以下依赖包:" -ForegroundColor Yellow
    $missingPackages | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host ""
    $install = Read-Host "是否现在安装? (y/n)"
    if ($install -eq 'y') {
        Write-Host "安装依赖..." -ForegroundColor Yellow
        pip install -r requirements.txt
    }
}

Write-Host ""
Write-Host "=" * 60 -ForegroundColor Cyan

# 启动Flask应用
python backend/app.py
