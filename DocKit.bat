@echo off
setlocal

:: Run DocKit from source without a console window
if "%~1"=="" (
    start "" pythonw.exe "%~dp0markitdown_gui.py"
) else (
    start "" pythonw.exe "%~dp0markitdown_gui.py" "%~1"
)
