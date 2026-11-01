@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Сначала запустите INSTALL.bat.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" analyze.py %*
set "ANALYSIS_EXIT=%errorlevel%"
if not "%ANALYSIS_EXIT%"=="0" echo Анализ не завершён. Прочитайте ошибку выше.
pause
exit /b %ANALYSIS_EXIT%
