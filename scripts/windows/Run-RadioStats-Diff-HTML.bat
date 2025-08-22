@echo off
setlocal ENABLEDELAYEDEXPANSION
set SCRIPT_DIR=%~dp0
set PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe

if "%~3"=="" (
  echo Usage: %~nx0 BEFORE_TXT AFTER_TXT OUTPUT_HTML
  exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Parse-ArubaRadioStats-Diff.ps1" ^
  -BeforeFile "%~1" -AfterFile "%~2" -OutputHtml "%~3"

exit /b %ERRORLEVEL%