#!/bin/bash
# Bruce Doc Converter - 自动检测并使用系统 Python

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 检测 Python 命令（优先 python3）
if command -v python3 &> /dev/null; then
    PYTHON_CMD=python3
elif command -v python &> /dev/null; then
    PYTHON_CMD=python
else
    echo "错误: 未找到 Python。请先安装 Python 3.6 或更高版本。"
    echo "  macOS:   brew install python3"
    echo "  Ubuntu:  sudo apt update && sudo apt install python3"
    echo "  其他:    https://www.python.org/downloads/"
    exit 1
fi

# 运行转换脚本（依赖会自动安装到用户目录）
exec "$PYTHON_CMD" "$SCRIPT_DIR/scripts/convert_document.py" "$@"
