@echo off
chcp 65001 >nul
echo.
echo ===============================================
echo  微信批量发送助手 CLI — Windows 安装程序
echo ===============================================
echo.

:: 检查 Python
python --version 2>nul
if errorlevel 1 (
    echo  [错误] 未找到 Python 3
    echo  请先安装: https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo  [OK] Python: %%v

:: 创建虚拟环境
if exist ".venv\Scripts\python.exe" (
    echo  [OK] 已有虚拟环境 .venv
) else (
    echo  [信息] 创建虚拟环境...
    python -m venv .venv
)

:: 安装基础依赖
echo  [信息] 安装基础依赖...
.venv\Scripts\pip install -q -r requirements.txt

:: 安装 Windows 自动化依赖
if exist "requirements-win.txt" (
    echo  [信息] 安装 Windows 自动化依赖...
    .venv\Scripts\pip install -q -r requirements-win.txt
)

echo.
echo ===============================================
echo  安装完成！
echo ===============================================
echo.
echo  快速上手:
echo   .venv\Scripts\python app\cli.py setup    # 首次配置
echo   .venv\Scripts\python app\cli.py config   # 查看配置
echo   .venv\Scripts\python app\cli.py status   # 查看任务
echo   .venv\Scripts\python app\cli.py send     # 立即发送
echo   .venv\Scripts\python app\cli.py daemon   # 守护进程
echo.
echo  详细文档: type README.md
echo ===============================================
pause
