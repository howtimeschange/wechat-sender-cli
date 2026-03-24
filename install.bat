@echo off
:: ================================================
:: 微信批量发送助手 — Windows 一键安装脚本
:: 用法: install.bat
:: ================================================
chcp 65001 >nul
echo.
echo ===============================================
echo  微信批量发送助手 CLI — Windows 安装程序
echo ===============================================
echo.

:: 检查 Python 3
python --version 2>nul
if errorlevel 1 (
    echo  [错误] 未找到 Python 3
    echo  请先安装 Python 3: https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo  [OK] Python 版本: %%v

:: 创建虚拟环境
if exist ".venv\Scripts\python.exe" (
    echo  [OK] 已有虚拟环境 .venv
) else (
    echo  [信息] 创建虚拟环境 .venv ...
    python -m venv .venv
)

:: 安装依赖
echo  [信息] 安装 Python 依赖 ...
.venv\Scripts\pip install -q -r requirements.txt
if errorlevel 1 (
    echo  [错误] 依赖安装失败
    pause
    exit /b 1
)
echo  [OK] 依赖安装完成

echo.
echo ===============================================
echo  安装完成！
echo ===============================================
echo.
echo  使用方法：
echo   .venv\Scripts\python app\cli.py setup    # 首次配置（交互向导）
echo   .venv\Scripts\python app\cli.py config   # 查看当前配置
echo   .venv\Scripts\python app\cli.py status   # 查看任务状态
echo   .venv\Scripts\python app\cli.py send     # 立即发送（需在 Mac 上）
echo   .venv\Scripts\python app\cli.py daemon   # 守护进程模式
echo.
echo  注意：Windows 版仅支持查看/配置功能
echo        微信自动化发送功能需要在 Mac 上运行
echo.
echo  详细文档：type README.md
echo ===============================================
pause
