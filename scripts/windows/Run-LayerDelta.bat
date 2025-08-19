@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem === Save as CP932 (Shift-JIS), CRLF, WITHOUT BOM ===
rem Run by double-click or from CMD (PowerShell から実行する場合は `cmd /c` を推奨)

rem この .bat のあるフォルダへ移動（UNC/OneDrive/日本語パス対応）
pushd "%~dp0" || (echo [ERROR] pushd failed & exit /b 1)

rem ---- 設定 ----
set "PS=Generate-TeamsNet-Report.ps1"
set "CSV=%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv"
rem targets.csv / floors.csv は 1 階層上
set "TARGETS=%CD%\..\targets.csv"
set "FLOORFILE=%CD%\..\floors.csv"
set "OUTDIR=%CD%\Output"
set "OUT=%OUTDIR%\TeamsNet-Report.xlsx"

rem ---- 事前チェック ----
if not exist "%PS%"       ( echo [ERROR] PS script not found: "%PS%" & popd & exit /b 1 )
if not exist "%TARGETS%"  ( echo [ERROR] targets.csv not found one level up: "%TARGETS%" & popd & exit /b 1 )
if not exist "%CSV%"      ( echo [ERROR] data CSV not found: "%CSV%" & popd & exit /b 1 )

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

rem ---- 実行（floors.csv があれば -FloorMap を付与）----
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

set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo [ERROR] PowerShell script failed. ERRORLEVEL=%RC%
  popd & exit /b %RC%
)

echo [OK] Report generated: "%OUT%"
popd
exit /b 0