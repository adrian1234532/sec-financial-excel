@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python not found: %PYTHON_EXE%
    exit /b 1
)
cd /d "%SCRIPT_DIR%"
if not "%~1"=="" goto arguments
echo SEC financial statements to Excel
set /p "TICKER=Ticker symbol, e.g. GOOGL: "
if "%TICKER%"=="" exit /b 1
"%PYTHON_EXE%" "%SCRIPT_DIR%sec_data_cli.py" --ticker %TICKER% --xlsx
exit /b %errorlevel%
:arguments
set "FIRST_ARG=%~1"
if "%FIRST_ARG:~0,1%"=="-" goto options
"%PYTHON_EXE%" "%SCRIPT_DIR%sec_data_cli.py" --ticker %* --xlsx
exit /b %errorlevel%
:options
"%PYTHON_EXE%" "%SCRIPT_DIR%sec_data_cli.py" %*
exit /b %errorlevel%
