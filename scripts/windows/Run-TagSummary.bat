@echo off
setlocal ENABLEDELAYEDEXPANSION

rem ===========================================
rem TagSummary レポート実行バッチ（PS 5.1 互換）
rem 使い方:
rem   Run-TagSummary.bat [CsvPath] [ThresholdMs] [TagColumn] [HostColumn] [LatencyColumn]
rem 例:
rem   Run-TagSummary.bat ".\teams_net_quality_merged.csv" 120 probe target rtt_ms
rem 引数未指定時はカレントの最新 *.csv を検索、しきい値は 100ms
rem ===========================================

set SCRIPT_DIR=%~dp0
set ROOT_DIR=%SCRIPT_DIR%\..
pushd "%ROOT_DIR%" >NUL

set CSV_PATH=%~1
set THRESHOLD=%~2
set TAGCOL=%~3
set HOSTCOL=%~4
set LATCOL=%~5

if "%THRESHOLD%"=="" set THRESHOLD=100

if "%CSV_PATH%"=="" (
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "*.csv" 2^>NUL') do (
    if not defined CSV_PATH set CSV_PATH=%%~fF
  )
)

if "%CSV_PATH%"=="" (
  echo [Run-TagSummary] ERROR: CSV が見つかりません。引数で CSV パスを指定するか、カレントに *.csv を置いてください。
  exit /b 1
)

set PS_SCRIPT=.\scripts\windows\Generate-TeamsNet-TagSummary.ps1
if not exist "%PS_SCRIPT%" (
  echo [Run-TagSummary] ERROR: %PS_SCRIPT% が見つかりません。
  exit /b 1
)

echo [Run-TagSummary] CSV: %CSV_PATH%
echo [Run-TagSummary] ThresholdMs: %THRESHOLD%
if not "%TAGCOL%"=="" echo [Run-TagSummary] TagColumn: %TAGCOL%
if not "%HOSTCOL%"=="" echo [Run-TagSummary] HostColumn: %HOSTCOL%
if not "%LATCOL%"=="" echo [Run-TagSummary] LatencyColumn: %LATCOL%

rem PowerShell 5.1（Windows付属）で実行
powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%PS_SCRIPT%" -CsvPath "%CSV_PATH%" -Output ".\Output" -ThresholdMs %THRESHOLD% ^
  %TAGCOL:=-TagColumn "%TAGCOL%" % ^
  %HOSTCOL:=-HostColumn "%HOSTCOL%" % ^
  %LATCOL:=-LatencyColumn "%LATCOL%" %

set ERR=%ERRORLEVEL%
if NOT "%ERR%"=="0" (
  echo [Run-TagSummary] ERRORLEVEL=%ERR%
  popd >NUL
  exit /b %ERR%
)

echo [Run-TagSummary] 完了: .\Output\TagSummary.xlsx / TagSummary.csv
popd >NUL
exit /b 0
