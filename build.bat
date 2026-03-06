@echo off
title Qobuz RPC - Build EXE
echo.
echo ========================================
echo   Qobuz RPC - Build Standalone EXE
echo ========================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Python not found.
    pause
    exit /b 1
)

echo [*] Installing PyInstaller...
python -m pip install pyinstaller
echo.

echo [*] Building GUI version (QobuzRPC.exe)...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "QobuzRPC" ^
    --icon "icon.ico" ^
    --add-data "icon.ico;." ^
    --add-data "icon.png;." ^
    --add-data "config.json;." ^
    --hidden-import "pystray._win32" ^
    qobuz_rpc.py

echo.
echo [*] Building CLI version (QobuzRPC-CLI.exe)...
python -m PyInstaller --noconfirm --onefile --console ^
    --name "QobuzRPC-CLI" ^
    --icon "icon.ico" ^
    qobuz_rpc_cli.py

echo.
echo ========================================
if exist "dist\QobuzRPC.exe" (
    echo [+] GUI:  dist\QobuzRPC.exe
) else (
    echo [!] GUI build failed
)
if exist "dist\QobuzRPC-CLI.exe" (
    echo [+] CLI:  dist\QobuzRPC-CLI.exe
) else (
    echo [!] CLI build failed
)
echo ========================================
echo.
echo Copy the .exe files along with config.json,
echo icon.ico, and icon.png to wherever you want.
echo.
pause
