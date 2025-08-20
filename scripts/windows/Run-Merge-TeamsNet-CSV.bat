@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 932 >nul 2>&1

rem ===================== 設定 =====================
rem PowerShell スクリプト名（同ディレクトリ想定）
set "SCRIPT=Merge-TeamsNet-CSV.ps1"

rem マージ対象の親フォルダをセミコロン(;)区切りで列挙（相対/絶対どちらも可）
rem 例1: set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"
rem 例2: set "FOLDERS=.\部署A\8F 東;.\部署B\10F 西"
rem ※ フォルダ名にスペース/日本語があっても OK。ここでは**各要素を引用符で囲まない**でください。
set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"

rem タグ（; 区切り）。空なら各 FOLDERS 要素の**末端フォルダ名**を既定タグとして使用します。
rem 例: set "TAGS=8F-A;10F-B"
set "TAGS="

rem 1=サブフォルダも再帰、0=直下のみ
set "RECURSE=1"

rem 出力先（単一ファイル）
set "OUTPUT=.\merged_teams_net_quality.csv"

rem ====== 日別分割の設定（SplitByDate を使う場合）======
rem 1=日別に分割して出力 / 0=単一ファイル出力
set "SPLIT_BY_DATE=1"
rem 日付判定に使う列名（CSV 内の列）
set "DATE_COLUMN=timestamp"
rem 出力ファイル名のフォーマット（例: merged_yyyyMMdd.csv）
set "DATE_FORMAT=yyyyMMdd"
rem 分割出力のベースディレクトリ
set "OUTPUT_DIR=.\merged_by_date"

rem ================ ここから処理 ================
set "CSV_LIST="
set "TAG_LIST="

rem -- セミコロン区切りの FOLDERS を安全に 1 件ずつ取り出して処理 --
set "REST=%FOLDERS%"
:__split_loop
if not defined REST goto __split_done
for /f "tokens=1* delims=; eol=§" %%A in ("%REST%") do (
  set "ITEM=%%~A"
  set "REST=%%~B"
)
if defined ITEM (
  call :__process_folder "!ITEM!"
  set "ITEM="
)
goto __split_loop
:__split_done

rem 明示 TAGS があれば置き換え（; 区切りのまま）
if not "%TAGS%"=="" set "TAG_LIST=%TAGS%"

if "%CSV_LIST%"=="" (
  echo [ERROR] path_hop_quality.csv が見つかりませんでした。>&2
  exit /b 1
)

rem PowerShell 実行（名前付き引数）
if "%SPLIT_BY_DATE%"=="1" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
    -InputCsvs "%CSV_LIST%" ^
    -Tags "%TAG_LIST%" ^
    -SplitByDate ^
    -DateColumn "%DATE_COLUMN%" ^
    -DateFormat "%DATE_FORMAT%" ^
    -OutputDir "%OUTPUT_DIR%"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
    -InputCsvs "%CSV_LIST%" ^
    -Tags "%TAG_LIST%" ^
    -Output "%OUTPUT%"
)

exit /b !ERRORLEVEL!

rem ------------------------------------------------
:__process_folder
rem %~1 = FOLDERS 要素（フォルダパス、引用符なし想定）
set "D=%~1"
if not defined D exit /b 0

if "%RECURSE%"=="1" (
  for /r "%D%" %%F in (path_hop_quality.csv) do (
    call :__append "%%~fF" "%~nx1"
  )
) else (
  if exist "%D%\path_hop_quality.csv" (
    call :__append "%D%\path_hop_quality.csv" "%~nx1"
  )
)
exit /b 0

:__append
rem %~1 = CSV のフルパス, %~2 = 既定タグ（FOLDERS 要素の末端名）
set "CSV=%~1"
set "DEF_TAG=%~2"

if "%CSV_LIST%"=="" (set "CSV_LIST=%CSV%") else (set "CSV_LIST=%CSV_LIST%;%CSV%")
if "%TAG_LIST%"=="" (set "TAG_LIST=%DEF_TAG%") else (set "TAG_LIST=%TAG_LIST%;%DEF_TAG%")
exit /b 0
