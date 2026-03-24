# ================================================
# 微信批量发送助手 — Windows PowerShell 一键安装脚本
# 用法 (管理员 PowerShell): irm https://.../install.ps1 | iex
# ================================================
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Repo = "howtimeschange/wechat-sender-cli"
$Temp = "$env:TEMP\wechat-sender-cli"

Write-Host "📦 微信批量发送助手 CLI — Windows 安装程序" -ForegroundColor Cyan
Write-Host "============================================"

# 检查 Python
try {
    $PyVer = python --version 2>&1
    Write-Host "✓ Python: $PyVer" -ForegroundColor Green
} catch {
    Write-Host "❌ 未找到 Python 3，请先安装 https://www.python.org/downloads/windows/" -ForegroundColor Red
    Write-Host "安装完成后重新运行此脚本" -ForegroundColor Yellow
    exit 1
}

# 下载仓库
if (Test-Path $Temp) { Remove-Item $Temp -Recurse -Force }
Write-Host "📥 克隆仓库..."
git clone --quiet "https://github.com/$Repo.git" $Temp

# 安装依赖
Set-Location $Temp
Write-Host "📦 安装 Python 依赖..."
if (-not (Test-Path ".venv")) {
    python -m venv .venv
}
.venv\Scripts\pip install -q -r requirements.txt
.venv\Scripts\pip install -q -r requirements-win.txt
Write-Host "✓ 依赖安装完成" -ForegroundColor Green

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "✅ 安装完成！" -ForegroundColor Green
Write-Host ""
Write-Host "快速上手：" -ForegroundColor White
Write-Host "  .venv\Scripts\python app\cli.py setup    # 首次配置（交互向导）" -ForegroundColor Yellow
Write-Host "  .venv\Scripts\python app\cli.py config   # 查看配置" -ForegroundColor Yellow
Write-Host "  .venv\Scripts\python app\cli.py status   # 查看任务" -ForegroundColor Yellow
Write-Host "  .venv\Scripts\python app\cli.py send     # 立即发送" -ForegroundColor Yellow
Write-Host ""
Write-Host "详细文档：type README.md" -ForegroundColor Gray
Write-Host "============================================"
