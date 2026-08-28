@echo off
set MOMO_HOME=%LOCALAPPDATA%\Momo
set PATH=%MOMO_HOME%\runtime\node;%PATH%
if not exist "%MOMO_HOME%\app\node_modules\electron\dist\electron.exe" (
  echo Momo is not installed yet. Run setup.ps1 first:
  echo   powershell -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
  pause
  exit /b 1
)
rem %~dp0 ends with a backslash which would escape the closing quote - append a dot
start "Momo" "%MOMO_HOME%\app\node_modules\electron\dist\electron.exe" "%~dp0."
