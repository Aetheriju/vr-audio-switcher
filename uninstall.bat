@echo off
title VR Audio Switcher - Uninstall
color 0C

echo.
echo  ========================================
echo   VR Audio Switcher - Uninstall
echo  ========================================
echo.
echo  This will remove shortcuts and config files.
echo  Program files will remain (delete this folder to fully remove).
echo.
pause

:: Remove shortcuts (check both OneDrive and plain Desktop)
del "%USERPROFILE%\Desktop\VR Audio Switcher.lnk" 2>nul
del "%USERPROFILE%\OneDrive\Desktop\VR Audio Switcher.lnk" 2>nul
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\VR Audio Switcher.lnk" 2>nul

:: Remove config/state files
del "%~dp0config.json" 2>nul
del "%~dp0vm_devices.json" 2>nul
del "%~dp0vm_state.json" 2>nul
del "%~dp0state.json" 2>nul
del "%~dp0presets.json" 2>nul
del "%~dp0switcher.log" 2>nul
del "%~dp0switcher.log.1" 2>nul
del "%~dp0wizard.log" 2>nul
del "%~dp0wizard.log.1" 2>nul
del "%~dp0_enum_apps.csv" 2>nul

echo.
echo  [OK] Shortcuts and config files removed.
echo  [OK] You can delete this folder to remove program files.
echo.
pause
