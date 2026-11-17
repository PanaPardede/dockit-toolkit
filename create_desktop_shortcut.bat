@echo off
setlocal

set "EXE=%~dp0dist\DocKit.exe"
set "SHORTCUT=%USERPROFILE%\Desktop\DocKit.lnk"

if not exist "%EXE%" (
    echo.
    echo dist\DocKit.exe was not found.
    echo Build it first:  python -m PyInstaller DocKit.spec --noconfirm
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$s = $ws.CreateShortcut('%SHORTCUT%'); " ^
  "$s.TargetPath = '%EXE%'; " ^
  "$s.WorkingDirectory = '%~dp0dist'; " ^
  "$s.IconLocation = '%EXE%,0'; " ^
  "$s.Description = 'DocKit - offline document toolkit'; " ^
  "$s.Save()"

echo.
echo Desktop shortcut created: %SHORTCUT%
echo.
pause
