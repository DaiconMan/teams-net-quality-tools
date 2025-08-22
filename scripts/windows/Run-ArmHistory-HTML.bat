@echo off
setlocal ENABLEDELAYEDEXPANSION
set SCRIPT_DIR=%~dp0
set PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe

if "%~2"=="" (
  echo Usage: %~nx0 ARM_HISTORY_TXT_OR_GLOB OUTPUT_HTML
  exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Parse-ArubaArmHistory.ps1" ^
  -InputFiles "%~1" -OutputHtml "%~2"

exit /b %ERRORLEVEL%