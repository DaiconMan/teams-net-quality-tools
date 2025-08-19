@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem === Save as CP932 (Shift-JIS), CRLF, WITHOUT BOM ===

rem Move to this script's folder (works with UNC/OneDrive/Japanese paths)
pushd "%~dp0" || (echo [ERROR] pushd failed & exit /b 1)

rem --- Settings ---
set "BASEDIR=%CD%"
set "PS=Generate-TeamsNet-Report.ps1"
set "CSV=%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv"
rem targets.csv / floors.csv are one level up from this .bat
set "TARGETS=%BASEDIR%\..\targets.csv"
set "FLOORFILE=%BASEDIR%\..\floors.csv"
set "OUTDIR=%BASEDIR%\Output"
set "OUT=%OUTDIR%\TeamsNet-Report.xlsx"

rem --- Pre checks ---
if not exist "%PS%" (
  echo [ERROR] PS script not found: "%PS%"
  goto :fail
)
if not exist "%TARGETS%" (
  echo [ERROR] targets.csv not found (expected one level up): "%TARGETS%"
  goto :fail
)
if not exist "%CSV%" (
  echo [ERROR] data CSV not found: "%CSV%"
  goto :fail
)

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

rem --- Run (with FloorMap if available) ---
if exist "%FLOORFILE%" (
  echo [INFO] Using FloorMap: "%FLOORFILE%"
  powershell -NoProfile -ExecutionPolicy Bypass ^
    -File "%PS%" ^
    -CsvPath "%CSV%" ^
    -TargetsCsv "%TARGETS%" ^
    -Output "%OUT%" ^
    -BucketMinutes 5 -ThresholdMs 100 ^
    -FloorMap "%FLOORFILE%" 2^>^&1
) else (
  echo [INFO] floors.csv not found one level up. Skipping floor coloring.
  powershell -NoProfile -ExecutionPolicy Bypass ^
    -File "%PS%" ^
    -CsvPath "%CSV%" ^
    -TargetsCsv "%TARGETS%" ^
    -Output "%OUT%" ^
    -BucketMinutes 5 -ThresholdMs 100 2^>^&1
)

set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
  echo [ERROR] PowerShell script failed. ERRORLEVEL=%ERR%
  goto :fail
)

echo [OK] Report generated: "%OUT%"
popd
exit /b 0

:fail
popd
exit /b %ERR%