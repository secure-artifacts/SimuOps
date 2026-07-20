@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%"

if not exist ".venv" (
    py -3 -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if not exist "app_config.json" (
    copy /Y "config.default.json" "app_config.json" >nul
)

if not exist "command_presets.json" (
    echo []> "command_presets.json"
)

if not exist "input_prefixes.json" (
    echo []> "input_prefixes.json"
)

echo.
echo Environment setup completed.
echo Run start.bat to launch the application.
pause
