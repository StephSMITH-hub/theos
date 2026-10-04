@echo off
title Theos - Convert Markdown to DOCX
cd /d "%~dp0"

:: 1. Verify if python is REALLY executable (not just WindowsApps 9009 stub)
python -c "import sys; exit(0)" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PY_CMD=python
    goto run_python
)

py -c "import sys; exit(0)" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PY_CMD=py
    goto run_python
)

:: 2. Check standard user AppData Python installations
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -c "exit(0)" >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PY_CMD="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        goto run_python
    )
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" -c "exit(0)" >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PY_CMD="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
        goto run_python
    )
)

:: 3. If Python is not yet configured, use the built-in native PowerShell engine
echo [Info] Working Python runtime not detected in system PATH.
echo [Info] Running conversion using built-in native engine...
echo.

if "%~1"=="" goto interactive_native
goto file_native

:interactive_native
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0convert_md_to_docx.ps1"
goto finish

:file_native
set "TARGET_FILE=%~f1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0convert_md_to_docx.ps1' -InputPath $env:TARGET_FILE"
goto finish

:run_python
echo [Running: %PY_CMD% "%~dp0convert_md_to_docx.py" %*]
%PY_CMD% "%~dp0convert_md_to_docx.py" %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Python finished with code %ERRORLEVEL%.
)
goto finish

:finish
echo.
pause
exit /b 0
