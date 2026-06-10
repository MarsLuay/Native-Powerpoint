@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%install-native-powerpoint-doc-editor.ps1"

if not exist "%PS_SCRIPT%" (
  echo Could not find installer script:
  echo   %PS_SCRIPT%
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*

if errorlevel 1 (
  echo.
  echo Native PowerPoint Doc Editor install failed.
  pause
  exit /b 1
)

echo.
echo Native PowerPoint Doc Editor install completed.
pause
