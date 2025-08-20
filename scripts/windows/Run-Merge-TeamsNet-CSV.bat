@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 932 >nul 2>&1

rem ===================== 設定 =====================
set "SCRIPT=Merge-TeamsNet-CSV.ps1"

rem セミコロン(;)区切りで“ディレクトリ”を列挙してください（スペース/日本語OK）
rem 例: set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"
set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"

rem タグ（; 区切り）。空なら各ディレクトリの末端名を自動採用します
rem 例: set "TAGS=8F-A;10F-B"
set "TAGS="

rem 1=サブフォルダも再帰、0=直下のみ
set "RECURSE=1"

rem 単一ファイル出力
set "OUTPUT=.\merged_teams_net_quality.csv"

rem ====== 日別分割（使う場合）======
set "SPLIT_BY_DATE=1"
set "DATE_COLUMN=timestamp"
set "DATE_FORMAT=yyyyMMdd"
set "OUTPUT_DIR=.\merged_by_date"

rem ================ ここから処理 ================
set "UNIT_LIST="
set "TAG_LIST="

rem -- セミコロン分割を安全に処理 --
set "REST=%FOLDERS%"
:__split_loop
if not defined REST goto __split_done
for /f "tokens=1* delims=; eol=§" %%A in ("%REST%") do (
  set "ITEM=%%~A"
  set "REST=%%~B"
)
if defined ITEM (
  call :__append_dir "!ITEM!"
  set "ITEM="
)
goto __split_loop
:__split_done

rem 明示 TAGS があれば置き換え
if not "%TAGS%"=="" set "TAG_LIST=%TAGS%"

if "%UNIT_LIST%"=="" (
  echo [ERROR] マージ対象ディレクトリが指定されていません。>&2
  exit /b 1
)

rem ---- PowerShell コマンドを安全に組み立て（可変引数を行末に置かない）----
set "PSCMD=powershell -NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%"" -InputCsvs ""%UNIT_LIST%"" -Tags ""%TAG_LIST%"""
if "%SPLIT_BY_DATE%"=="1" (
  set "PSCMD=!PSCMD! -SplitByDate -DateColumn ""%DATE_COLUMN%"" -DateFormat ""%DATE_FORMAT%"" -OutputDir ""%OUTPUT_DIR%"""
) else (
  set "PSCMD=!PSCMD! -Output ""%OUTPUT%"""
)
if "%RECURSE%"=="1" (
  set "PSCMD=!PSCMD! -Recurse"
)

rem 実行（行継続 ^ を使わず、そのまま実行）
%PSCMD%
exit /b !ERRORLEVEL!

:__append_dir
rem %~1 = ディレクトリパス（引用符なし）
set "D=%~1"
if not defined D exit /b 0

if "%UNIT_LIST%"=="" (set "UNIT_LIST=%D%") else (set "UNIT_LIST=%UNIT_LIST%;%D%")
if "%TAG_LIST%"=="" (set "TAG_LIST=%~nx1") else (set "TAG_LIST=%TAG_LIST%;%~nx1")
exit /b 0
