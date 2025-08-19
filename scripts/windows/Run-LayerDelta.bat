@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem === Save as CP932 (Shift-JIS), CRLF, WITHOUT BOM ===

rem この .bat のあるフォルダへ移動（UNC/OneDrive/日本語パス対応）
pushd "%~dp0" || (echo [ERROR] pushd failed & exit /b 1)

rem ---- 設定 ----
set "BASEDIR=%CD%"
set "PS=%BASEDIR%\Generate-TeamsNet-Report.ps1"
set "CSV=%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv"
rem targets.csv / floors.csv は 1 階層上
set "TARGETS=%BASEDIR%\..\targets.csv"
set "FLOORFILE=%BASEDIR%\..\floors.csv"
set "OUTDIR=%BASEDIR%\Output"
set "OUT=%OUTDIR%\TeamsNet-Report.xlsx"
set "LOG=%OUTDIR%\ps-error.log"

rem ---- 事前チェック ----
if not exist "%PS%"      ( echo [ERROR] PS script not found: "%PS%" & popd & exit /b 1 )
if not exist "%TARGETS%" ( echo [ERROR] targets.csv not found one level up: "%TARGETS%" & popd & exit /b 1 )
if not exist "%CSV%"     ( echo [ERROR] data CSV not found: "%CSV%" & popd & exit /b 1 )

if not exist "%OUTDIR%" mkdir "%OUTDIR%"
if exist "%LOG%" del "%LOG%" >nul 2>&1

rem ---- 引数組み立て（引用は最小限）----
set "ARGS=-CsvPath ""%CSV%"" -TargetsCsv ""%TARGETS%"" -Output ""%OUT%"" -BucketMinutes 5 -ThresholdMs 100"
if exist "%FLOORFILE%" (
  echo [INFO] Using FloorMap: "%FLOORFILE%"
  set "ARGS=%ARGS% -FloorMap ""%FLOORFILE%"""
) else (
  echo [INFO] floors.csv not found one level up. Skipping floor coloring.
)

rem ---- 実行（PowerShell側で try/catch、詳細を ps-error.log に出力）----
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; $VerbosePreference='Continue'; $WarningPreference='Continue';" ^
  "try { & '%PS%' %ARGS% } catch { " ^
  "  '--- PowerShell Error (Generate-TeamsNet-Report.ps1) ---' | Tee-Object -File '%LOG%';" ^
  "  $_ | Format-List * -Force | Out-String | Tee-Object -File '%LOG%' -Append;" ^
  "  if ($_.InvocationInfo) { '--- InvocationInfo ---' | Tee-Object -File '%LOG%' -Append; $_.InvocationInfo | Format-List * -Force | Out-String | Tee-Object -File '%LOG%' -Append };" ^
  "  if ($_.ScriptStackTrace) { '--- ScriptStackTrace ---' | Tee-Object -File '%LOG%' -Append; ($_.ScriptStackTrace) | Tee-Object -File '%LOG%' -Append };" ^
  "  if ($Error.Count -gt 0) { '--- $Error[0] ---' | Tee-Object -File '%LOG%' -Append; ($Error[0] | Format-List * -Force | Out-String) | Tee-Object -File '%LOG%' -Append };" ^
  "  exit 1 }"

set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo [ERROR] PowerShell script failed. ERRORLEVEL=%RC%
  echo [HINT] See detailed log: "%LOG%"
  popd & exit /b %RC%
)

echo [OK] Report generated: "%OUT%"
popd
exit /b 0