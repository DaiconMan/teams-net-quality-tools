@echo off
setlocal EnableExtensions

rem ===== 設定（この bat と同じフォルダに PS1／Output を置く想定。targets/floors は 1 階層上）=====
set "SCRIPT=Generate-TeamsNet-Report.ps1"
set "CSV=%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv"
set "TARGETS=..\targets.csv"
set "FLOORS=..\floors.csv"
set "OUTDIR=Output"
set "OUTFILE=TeamsNet-Report.xlsx"

rem ── 集計パラメータ
set "BUCKET_MIN=5"
set "THRESHOLD_MS=100"

set "VERBOSE=0"     rem 1=詳細表示ON（PowerShellの -Verbose）
set "DEBUG_HOLD=1"  rem 1=終了時に停止
rem ===============================================================

rem 文字コードを CP932 に（日本語表示の乱れ防止）
chcp 932 >nul 2>&1

rem 作業フォルダへ移動（OneDrive/UNC/日本語パス対応）
set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] 作業フォルダへ移動できません: "%BASE%"
  if "%DEBUG_HOLD%"=="1" pause
  exit /b 1
)

rem 拡張子補正
for %%Z in ("%SCRIPT%") do set "EXT=%%~xZ"
if /I not "%EXT%"==".ps1" set "SCRIPT=%SCRIPT%.ps1"

rem 絶対パス化
for %%I in ("%SCRIPT%")   do set "ABS_SCRIPT=%%~fI"
for %%I in ("%CSV%")      do set "ABS_CSV=%%~fI"
for %%I in ("%TARGETS%")  do set "ABS_TARGETS=%%~fI"
for %%I in ("%FLOORS%")   do set "ABS_FLOORS=%%~fI"
for %%I in ("%OUTDIR%")   do set "ABS_OUTDIR=%%~fI"
for %%I in ("%OUTFILE%")  do set "ABS_OUT=%%~fI"

rem 存在チェック
if not exist "%ABS_SCRIPT%"  ( echo [ERROR] スクリプトが見つかりません: "%ABS_SCRIPT%" & goto :fail )
if not exist "%ABS_CSV%"     ( echo [ERROR] データCSVが見つかりません: "%ABS_CSV%" & goto :fail )
if not exist "%ABS_TARGETS%" ( echo [ERROR] targets.csv が見つかりません（1階層上想定）: "%ABS_TARGETS%" & goto :fail )
if not exist "%ABS_OUTDIR%"  ( mkdir "%ABS_OUTDIR%" >nul 2>&1 )

rem 使用する PowerShell を選択（あれば pwsh、なければ powershell）
set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

set "EXTRA="
if "%VERBOSE%"=="1" set "EXTRA=-Verbose"

echo * 実行開始: %date% %time%
echo   PS         : %PS%
echo   Script     : %ABS_SCRIPT%
echo   CsvPath    : %ABS_CSV%
echo   TargetsCsv : %ABS_TARGETS%
if exist "%ABS_FLOORS%" (
  echo   FloorMap   : %ABS_FLOORS%
) else (
  echo   FloorMap   : (なし)
)
echo   Output     : %ABS_OUT%
echo   Bucket(min): %BUCKET_MIN%
echo   Threshold  : %THRESHOLD_MS% ms
if "%VERBOSE%"=="1" echo   Verbose    : on

echo --- 実行コマンド ---
if exist "%ABS_FLOORS%" (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -CsvPath "%ABS_CSV%" -TargetsCsv "%ABS_TARGETS%" -FloorMap "%ABS_FLOORS%" -Output "%ABS_OUT%" -BucketMinutes %BUCKET_MIN% -ThresholdMs %THRESHOLD_MS% %EXTRA%
) else (
  echo "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -CsvPath "%ABS_CSV%" -TargetsCsv "%ABS_TARGETS%" -Output "%ABS_OUT%" -BucketMinutes %BUCKET_MIN% -ThresholdMs %THRESHOLD_MS% %EXTRA%
)
echo --------------------

rem 実行
if exist "%ABS_FLOORS%" (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -CsvPath "%ABS_CSV%" -TargetsCsv "%ABS_TARGETS%" -FloorMap "%ABS_FLOORS%" -Output "%ABS_OUT%" -BucketMinutes %BUCKET_MIN% -ThresholdMs %THRESHOLD_MS% %EXTRA%
) else (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ABS_SCRIPT%" -CsvPath "%ABS_CSV%" -TargetsCsv "%ABS_TARGETS%" -Output "%ABS_OUT%" -BucketMinutes %BUCKET_MIN% -ThresholdMs %THRESHOLD_MS% %EXTRA%
)

set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] 生成に失敗しました。上のメッセージを確認してください.
goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
if "%DEBUG_HOLD%"=="1" pause
exit /b %RC%