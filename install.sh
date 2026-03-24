#!/bin/bash
# ================================================
# 微信批量发送助手 — macOS 安装脚本
# 用法: bash install.sh
# ================================================
set -e

cd "$(dirname "$0")"

echo "🍎 微信批量发送助手 CLI — macOS 安装程序"
echo "============================================"

# 检查 Python 3
if ! command -v python3 &>/dev/null; then
    echo "❌ 未找到 python3，请先安装 Python 3"
    echo "   官网: https://www.python.org/downloads/macos/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "✓ Python 版本: $PYTHON_VERSION"

# 创建/激活虚拟环境
if [ -d ".venv" ]; then
    echo "✓ 已有虚拟环境 .venv"
else
    echo "📦 创建虚拟环境..."
    python3 -m venv .venv
    echo "✓ 虚拟环境已创建"
fi

# 安装基础依赖
echo "📦 安装基础依赖..."
.venv/bin/pip install -q -r requirements.txt
echo "✓ 基础依赖安装完成"

# 检查 osascript
if command -v osascript &>/dev/null; then
    echo "✓ osascript 可用（支持微信自动化）"
else
    echo "⚠️  未找到 osascript，无法使用自动化功能"
fi

echo ""
echo "============================================"
echo "✅ 安装完成！"
echo ""
echo "快速上手："
echo "  .venv/bin/python app/cli.py setup    # 首次配置"
echo "  .venv/bin/python app/cli.py config   # 查看配置"
echo "  .venv/bin/python app/cli.py status   # 查看任务"
echo "  .venv/bin/python app/cli.py send     # 立即发送"
echo "  .venv/bin/python app/cli.py daemon   # 守护进程"
echo ""
echo "详细文档：cat README.md"
echo "============================================"
