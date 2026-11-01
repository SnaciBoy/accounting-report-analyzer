@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto dependencies
if defined PYTHON_EXE goto selected
py -3 -c "import sys, venv; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 (
    py -3 -m venv .venv
    goto created
)
set "PYTHON_EXE=python"
:selected
"%PYTHON_EXE%" -c "import sys, venv; assert sys.version_info >= (3,10)" >nul 2>nul
if errorlevel 1 goto failed
"%PYTHON_EXE%" -m venv .venv
:created
if errorlevel 1 goto failed
:dependencies
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Установка завершена. Запустите START.bat.
pause
exit /b 0
:failed
echo Установка не завершена. Проверьте сообщение выше.
echo Можно задать путь: set "PYTHON_EXE=полный путь к python.exe"
pause
exit /b 1
