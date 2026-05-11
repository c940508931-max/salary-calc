# 薪资计算工具 - Windows 打包脚本
# 在 Windows 上以管理员身份运行 PowerShell，执行此脚本

Write-Host "🚀 路易小姐薪资计算工具 - Windows 打包开始" -ForegroundColor Cyan

# 1. 检查 Python
try {
    $pyVersion = python --version
    Write-Host "✅ Python: $pyVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ 未找到 Python，请先安装 Python 3.8+" -ForegroundColor Red
    Write-Host "   下载地址：https://www.python.org/downloads/windows/"
    Write-Host "   安装时记得勾选 'Add Python to PATH'"
    exit 1
}

# 2. 安装依赖
Write-Host "`n📦 安装依赖..." -ForegroundColor Yellow
pip install flask openpyxl pyinstaller
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 依赖安装失败" -ForegroundColor Red
    exit 1
}
Write-Host "✅ 依赖安装完成" -ForegroundColor Green

# 3. 获取脚本所在目录
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# 4. 打包
Write-Host "`n📦 正在打包..." -ForegroundColor Yellow
Remove-Item -Path "dist", "build", "*.spec" -Recurse -Force -ErrorAction SilentlyContinue

pyinstaller --onefile --windowed `
    --name "薪资计算工具" `
    --add-data "templates;templates" `
    --icon "$ScriptDir\app.ico" `
    --noconfirm `
    app.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✅ 打包成功！" -ForegroundColor Green
    Write-Host "   可执行文件在：$ScriptDir\dist\薪资计算工具.exe" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "📋 使用说明：" -ForegroundColor Yellow
    Write-Host "   双击 薪资计算工具.exe 运行" -ForegroundColor White
    Write-Host "   浏览器打开 http://localhost:5001" -ForegroundColor White
    Write-Host ""
    Write-Host "⚠️ 首次运行可能被 Windows Defender 拦截，点击"更多信息"→"仍要运行"" -ForegroundColor Magenta
} else {
    Write-Host "❌ 打包失败" -ForegroundColor Red
}
