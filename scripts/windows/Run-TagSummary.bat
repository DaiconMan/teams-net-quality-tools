@echo off
setlocal ENABLEDELAYEDEXPANSION

rem ===========================================
rem TagSummary レポート実行バッチ（PS 5.1 互換）
rem 使い方:
rem   Run-TagSummary.bat [CsvPath] [ThresholdMs] [TagColumn] [TargetColumn] [LatencyColumn]
rem 例:
rem   Run-TagSummary.bat ".\teams_net_quality_merged.csv" 120 probe target rtt_ms
rem 引数未指定時はカレントの最新 *.csv を検索、しきい値は 100ms
rem ===========================================

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%\.." >NUL

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
  echo [Run-TagSummary] ERROR: CSV が見つかりません。引数で CSV パスを指定するか、カレントに *.csv を置いてください。
  popd >NUL
  exit /b 1
)

set "PS_SCRIPT=.\scripts\windows\Generate-TeamsNet-TagSummary.ps1"
if not exist "%PS_SCRIPT%" (
  echo [Run-TagSummary] ERROR: %PS_SCRIPT% が見つかりません。
  popd >NUL
  exit /b 1
)

echo [Run-TagSummary] CSV: %CSV_PATH%
echo [Run-TagSummary] ThresholdMs: %THRESHOLD%
if not "%TAGCOL%"=="" echo [Run-TagSummary] TagColumn: %TAGCOL%
if not "%TGTCOL%"=="" echo [Run-TagSummary] TargetColumn: %TGTCOL%
if not "%LATCOL%"=="" echo [Run-TagSummary] LatencyColumn: %LATCOL%

rem オプション引数を安全に組み立て
set "OPT_TAG="
set "OPT_TGT="
set "OPT_LAT="
if not "%TAGCOL%"=="" set "OPT_TAG=-TagColumn \"%TAGCOL%\""
if not "%TGTCOL%"=="" set "OPT_TGT=-TargetColumn \"%TGTCOL%\""
if not "%LATCOL%"=="" set "OPT_LAT=-LatencyColumn \"%LATCOL%\""

powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%PS_SCRIPT%" -CsvPath "%CSV_PATH%" -Output ".\Output" -ThresholdMs %THRESHOLD% ^
  %OPT_TAG% %OPT_TGT% %OPT_LAT%

set "ERR=%ERRORLEVEL%"
if NOT "%ERR%"=="0" (
  echo [Run-TagSummary] ERRORLEVEL=%ERR%
  popd >NUL
  exit /b %ERR%
)

echo [Run-TagSummary] 完了: .\Output\TagSummary.xlsx / TagSummary.csv
popd >NUL
exit /b 0
