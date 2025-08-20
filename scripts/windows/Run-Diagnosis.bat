@echo off
setlocal ENABLEDELAYEDEXPANSION
rem ===========================================
rem TeamsNet Diagnosis runner (PS 5.1 / OneDrive・日本語パス対応)
rem 使い方:
rem   Run-Diagnosis.bat [CsvPath] [ThresholdMs] [BucketMinutes] [FloorMap]
rem 例:
rem   Run-Diagnosis.bat ".\Output\teams_net_quality_merged.csv" 100 60 ".\floors.csv"
rem ===========================================

set "SCRIPT_DIR=%~dp0"
for %%A in ("%SCRIPT_DIR%\..\..") do set "REPO_DIR=%%~fA"
pushd "%REPO_DIR%" >NUL

set "CSV_PATH=%~1"
set "THRESHOLD=%~2"
set "BUCKET=%~3"
set "FLOORS=%~4"

if "%THRESHOLD%"=="" set "THRESHOLD=100"
if "%BUCKET%"=="" set "BUCKET=60"

if "%CSV_PATH%"=="" (
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "*.csv" 2^>NUL') do (
    if not defined CSV_PATH set "CSV_PATH=%%~fF"
  )
)

if "%CSV_PATH%"=="" (
  echo [Run-Diagnosis] ERROR: CSV が見つかりません。
  popd >NUL
  exit /b 1
)

set "PS_SCRIPT=%SCRIPT_DIR%\Generate-TeamsNet-Diagnosis.ps1"
if not exist "%PS_SCRIPT%" (
  echo [Run-Diagnosis] ERROR: %PS_SCRIPT% が見つかりません。
  popd >NUL
  exit /b 1
)

set "OUT_DIR=%REPO_DIR%\Output"

echo [Run-Diagnosis] CSV: %CSV_PATH%
echo [Run-Diagnosis] ThresholdMs: %THRESHOLD%
echo [Run-Diagnosis] BucketMinutes: %BUCKET%
if not "%FLOORS%"=="" echo [Run-Diagnosis] FloorMap: %FLOORS%
echo [Run-Diagnosis] Output: %OUT_DIR%

set "OPT_FLOOR="
if not "%FLOORS%"=="" set "OPT_FLOOR=-FloorMap \"%FLOORS%\""

powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%PS_SCRIPT%" -CsvPath "%CSV_PATH%" -Output "%OUT_DIR%" -ThresholdMs %THRESHOLD% -BucketMinutes %BUCKET% ^
  %OPT_FLOOR%

set "ERR=%ERRORLEVEL%"
if NOT "%ERR%"=="0" (
  echo [Run-Diagnosis] ERRORLEVEL=%ERR%
  popd >NUL
  exit /b %ERR%
)

echo [Run-Diagnosis] 完了: %OUT_DIR%\TeamsNet_Diagnosis.xlsx
echo [Run-Diagnosis]       %OUT_DIR%\Diag_ByArea_AP_Target.csv
echo [Run-Diagnosis]       %OUT_DIR%\Diag_TimeOfDay.csv
popd >NUL
exit /b 0
