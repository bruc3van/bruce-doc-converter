# Bruce Doc Converter - PowerShell Edition
# Auto-detect and use system Python with UTF-8 output

# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Detect Python command
$PythonCmd = $null

if (Get-Command python -ErrorAction SilentlyContinue) {
    $PythonCmd = "python"
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $PythonCmd = "py"
} else {
    Write-Host "Error: Python not found. Please install Python 3.6 or higher." -ForegroundColor Red
    Write-Host "  Visit: https://www.python.org/downloads/"
    Write-Host "  Check 'Add Python to PATH' during installation"
    exit 1
}

# Build path to conversion script
$ScriptsDir = Join-Path $ScriptDir "scripts"
$ConvertScript = Join-Path $ScriptsDir "convert_document.py"

# Run conversion script (dependencies auto-install to user directory)
& $PythonCmd $ConvertScript $args
exit $LASTEXITCODE
