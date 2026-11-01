@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "APP_PYTHON="
set "APP_PYTHON_ARGS="
if exist ".venv\Scripts\python.exe" goto dependencies
if exist ".venv" goto broken
if defined PYTHON_EXE (
    call :try_python "%PYTHON_EXE%"
    if not defined APP_PYTHON goto missing
    goto create
)
call :try_python py -3
if defined APP_PYTHON goto create
for /f "delims=" %%P in ('where python 2^>nul') do call :try_python "%%P"
if defined APP_PYTHON goto create
for /f "delims=" %%P in ('where python3 2^>nul') do call :try_python "%%P"
if not defined APP_PYTHON goto missing
:create
echo Выбранный интерпретатор:
"%APP_PYTHON%" %APP_PYTHON_ARGS% -c "import sys; print(sys.executable); print(sys.version)"
echo Создание локального окружения Python...
"%APP_PYTHON%" %APP_PYTHON_ARGS% -m venv .venv
if errorlevel 1 goto failed
if not exist ".venv\Scripts\python.exe" goto broken
:dependencies
".venv\Scripts\python.exe" -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if errorlevel 1 goto broken
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Готово. Для анализа запустите START.bat.
pause
exit /b 0
:missing
echo Не найден подходящий Windows Python 3.10+ с модулями venv и ensurepip.
echo Установите Python с python.org, включив pip и добавление в PATH.
echo Если Python уже установлен, откройте терминал заново или задайте PYTHON_EXE.
goto failed
:broken
echo Локальное окружение .venv недоступно или устарело.
echo Переименуйте папку .venv и повторите INSTALL.bat.
:failed
echo Установка не завершена. Прочитайте сообщение выше.
pause
exit /b 1
:try_python
if defined APP_PYTHON exit /b 0
"%~1" %~2 -c "import sys, venv, ensurepip, sysconfig, ntpath; assert sys.version_info >= (3,10); assert ntpath.basename(sysconfig.get_path('scripts')).lower() == 'scripts'" >nul 2>nul
if errorlevel 1 exit /b 0
set "APP_PYTHON=%~1"
set "APP_PYTHON_ARGS=%~2"
exit /b 0
