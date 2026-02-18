@echo off
title VR Audio Switcher - Installer
color 0A

echo.
echo  ========================================
echo   VR Audio Switcher - Installer
echo  ========================================
echo.

:: Check if Python is already installed
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo  [OK] Python found
    goto :run_wizard
)

:: Try py launcher (Windows Store Python)
py -3 --version >nul 2>&1
if %errorlevel% equ 0 (
    echo  [OK] Python found (py launcher)
    goto :run_wizard_py
)

:: Python not found - download and install it
echo  [..] Python not found - downloading installer...
echo.

:: Download Python installer
set PYTHON_URL=https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe
set INSTALLER=%~dp0_python_installer.exe

powershell -Command "Write-Host '  Downloading Python 3.13...' ; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 ; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALLER%' -UseBasicParsing"

if not exist "%INSTALLER%" (
    echo.
    echo  [ERROR] Download failed. Please install Python manually:
    echo          https://www.python.org/downloads/
    echo.
    echo  Make sure to check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)

echo.
echo  [..] Installing Python (this may take a minute)...
echo      - Adding to PATH automatically
echo      - Installing pip
echo.

:: Install Python silently with PATH enabled
"%INSTALLER%" /passive InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1

if %errorlevel% neq 0 (
    echo.
    echo  [..] Silent install didn't work - launching interactive installer...
    echo      IMPORTANT: Check "Add Python to PATH" at the bottom!
    echo.
    "%INSTALLER%" InstallAllUsers=0 PrependPath=1
)

:: Clean up installer
del "%INSTALLER%" >nul 2>&1

:: Refresh PATH — find the newest Python install dynamically
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    set "PATH=%%D;%%D\Scripts;%PATH%"
)

:: Verify Python is now available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Python installation may require a restart.
    echo          Please close this window, restart your PC, then
    echo          double-click install.bat again.
    echo.
    pause
    exit /b 1
)

echo  [OK] Python installed successfully
echo.

:run_wizard
echo  [..] Starting setup wizard...
echo.
python "%~dp0setup_wizard.py"
goto :done

:run_wizard_py
echo  [..] Starting setup wizard...
echo.
py -3 "%~dp0setup_wizard.py"
goto :done

:done
if %errorlevel% neq 0 (
    echo.
    echo  Something went wrong. Check the error above.
    pause
) else (
    echo.
    echo  Setup wizard launched successfully.
    echo  You can close this window.
    timeout /t 5 >nul
)
