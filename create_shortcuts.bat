@echo off
setlocal

:: Adds DocKit to the Start Menu (so Windows Search finds it) and to the Desktop.
set "EXE=%~dp0dist\DocKit.exe"
set "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs\DocKit.lnk"

if not exist "%EXE%" (
    echo.
    echo dist\DocKit.exe was not found.
    echo Build it first:  python -m PyInstaller DocKit.spec --noconfirm
    echo.
    pause
    exit /b 1
)

:: Desktop path from Windows (handles OneDrive-redirected desktops)
powershell -NoProfile -Command ^
  "$exe = '%EXE%';" ^
  "$targets = @('%STARTMENU%', (Join-Path ([Environment]::GetFolderPath('Desktop')) 'DocKit.lnk'));" ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "foreach ($t in $targets) {" ^
  "  $s = $ws.CreateShortcut($t); $s.TargetPath = $exe;" ^
  "  $s.WorkingDirectory = (Split-Path $exe); $s.IconLocation = \"$exe,0\";" ^
  "  $s.Description = 'DocKit - offline document toolkit'; $s.Save();" ^
  "  Write-Host ('Created: ' + $t) }"

echo.
echo Search "DocKit" in the Start menu to open it. Right-click it there to pin it.
echo.
pause
