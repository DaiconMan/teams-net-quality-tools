@echo off
setlocal ENABLEDELAYEDEXPANSION

rem ===========================================
rem TagSummary レポート実行バッチ（PS 5.1 互換 / OneDrive・日本語パス対応）
rem 使い方:
rem   Run-TagSummary.bat [CsvPath] [ThresholdMs] [TagColumn] [TargetColumn] [LatencyColumn]
rem 例:
rem   Run-TagSummary.bat ".\Output\teams_net_quality_merged.csv" 120 probe target rtt_ms
rem ===========================================

set "SCRIPT_DIR=%~dp0"
rem リポジトリ直下（... \scripts\windows の2つ上）
for %%A in ("%SCRIPT_DIR%\..\..") do set "REPO_DIR=%%~fA"
pushd "%REPO_DIR%" >NUL

set "CSV_PATH=%~1"
set "THRESHOLD=%~2"
set "TAGCOL=%~3"
set "TGTCOL=%~4"
set "LATCOL=%~5"

if "%THRESHOLD%"=="" set "THRESHOLD=100"

if "%CSV_PATH%"=="" (
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "*.csv" 2^>NUL') do (
    if not defined CSV_PATH set "CSV_PATH=%%~fF"
  )
)

if "%CSV_PATH%"=="" (
  echo [Run-TagSummary] ERROR: CSV が見つかりません。引数で CSV パスを指定するか、直下に *.csv を置いてください。
  popd >NUL
  exit /b 1
)

set "PS_SCRIPT=%SCRIPT_DIR%\Generate-TeamsNet-TagSummary.ps1"
if not exist "%PS_SCRIPT%" (
  echo [Run-TagSummary] ERROR: %PS_SCRIPT% が見つかりません。
  popd >NUL
  exit /b 1
)

rem 絶対パスの Output（必ず作られる）
set "OUT_DIR=%REPO_DIR%\Output"

echo [Run-TagSummary] CSV: %CSV_PATH%
echo [Run-TagSummary] ThresholdMs: %THRESHOLD%
echo [Run-TagSummary] Output: %OUT_DIR%
if not "%TAGCOL%"=="" echo [Run-TagSummary] TagColumn: %TAGCOL%
if not "%TGTCOL%"=="" echo [Run-TagSummary] TargetColumn: %TGTCOL%
if not "%LATCOL%"=="" echo [Run-TagSummary] LatencyColumn: %LATCOL%

set "OPT_TAG=" & set "OPT_TGT=" & set "OPT_LAT="
if not "%TAGCOL%"=="" set "OPT_TAG=-TagColumn \"%TAGCOL%\""
if not "%TGTCOL%"=="" set "OPT_TGT=-TargetColumn \"%TGTCOL%\""
if not "%LATCOL%"=="" set "OPT_LAT=-LatencyColumn \"%LATCOL%\""

powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%PS_SCRIPT%" -CsvPath "%CSV_PATH%" -Output "%OUT_DIR%" -ThresholdMs %THRESHOLD% ^
  %OPT_TAG% %OPT_TGT% %OPT_LAT%

set "ERR=%ERRORLEVEL%"
if NOT "%ERR%"=="0" (
  echo [Run-TagSummary] ERRORLEVEL=%ERR%
  popd >NUL
  exit /b %ERR%
)

echo [Run-TagSummary] 完了: %OUT_DIR%\TagSummary.xlsx / TagSummary.csv
popd >NUL
exit /b 0
