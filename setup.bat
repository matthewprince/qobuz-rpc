@echo off
title Qobuz Discord RPC - Setup
echo.
echo ========================================
echo   Qobuz Discord RPC - First Time Setup
echo ========================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Python not found. Install Python 3.10+ from python.org
    echo     Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo [*] Installing dependencies...
pip install -r requirements.txt
echo.

if %errorlevel% neq 0 (
    echo [!] Dependency install failed. Check errors above.
    pause
    exit /b 1
)

echo [+] Setup complete!
echo.
echo Next steps:
echo   1. Create a Discord app at https://discord.com/developers/applications
echo   2. Copy the Application ID into config.json
echo   3. (Optional) Upload art assets - see README.md
echo   4. Run start.bat to launch
echo.
pause
