@echo off
REM Bruce Doc Converter - 自动检测并使用系统 Python

REM 获取脚本所在目录
set "SCRIPT_DIR=%~dp0"

REM 检测 Python 命令
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
    goto :found
)

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=py"
    goto :found
)

echo 错误: 未找到 Python。请先安装 Python 3.6 或更高版本。
echo   访问: https://www.python.org/downloads/
echo   安装时请勾选 "Add Python to PATH"
exit /b 1

:found
REM 运行转换脚本（依赖会自动安装到用户目录）
"%PYTHON_CMD%" "%SCRIPT_DIR%scripts\convert_document.py" %*
