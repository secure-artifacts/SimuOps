@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%"

if not exist ".venv\Scripts\pythonw.exe" (
    echo Missing virtual environment. Run setup.bat first.
    pause
    exit /b 1
)

if not exist "app_config.json" (
    copy /Y "config.default.json" "app_config.json" >nul
)

if not exist "command_presets.json" (
    echo []> "command_presets.json"
)

if not exist "input_prefixes.json" (
    echo []> "input_prefixes.json"
)

start "" ".venv\Scripts\pythonw.exe" "main.pyw"
