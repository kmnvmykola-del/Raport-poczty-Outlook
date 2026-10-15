@echo off
chcp 65001 >nul
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Policz-wiadomosci.ps1" -Menu
set "kodZakonczenia=%ERRORLEVEL%"
echo.
if not "%kodZakonczenia%"=="0" echo Nie udało się utworzyć raportu. Przeczytaj komunikat powyżej i plik Instrukcja.md.
echo Naciśnij dowolny klawisz, aby zamknąć okno.
pause >nul
exit /b %kodZakonczenia%
