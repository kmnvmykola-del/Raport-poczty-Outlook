@echo off
chcp 65001 >nul
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Policz-wiadomosci.ps1" -PokazSkrzynki -Menu
set "kodZakonczenia=%ERRORLEVEL%"
echo.
echo Naciśnij dowolny klawisz, aby zamknąć okno.
pause >nul
exit /b %kodZakonczenia%
