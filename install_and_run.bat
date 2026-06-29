@echo off
echo ============================================
echo   Salary Slip Generator - Installer
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please download Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/3] Python found. Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

IF ERRORLEVEL 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/3] Creating desktop shortcut...
set SCRIPT_DIR=%~dp0
set SHORTCUT_TARGET=%SCRIPT_DIR%run.bat
set DESKTOP=%USERPROFILE%\Desktop

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\Salary Slip Generator.lnk'); $s.TargetPath = '%SHORTCUT_TARGET%'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.IconLocation = 'shell32.dll,1'; $s.Description = 'Salary Slip Generator'; $s.Save()"

echo [3/3] Done!
echo.
echo ============================================
echo   Installation Complete!
echo   Default password: admin123
echo   (Change it in Settings after first login)
echo ============================================
echo.
echo Launching application...
python salary_slip_app.py
pause
