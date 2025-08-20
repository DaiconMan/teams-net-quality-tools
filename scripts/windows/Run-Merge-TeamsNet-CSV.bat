@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 932 >nul 2>&1

rem ===================== 設定 =====================
rem PowerShell スクリプト名（同ディレクトリ想定）
set "SCRIPT=Merge-TeamsNet-CSV.ps1"

rem マージ対象の親フォルダ（; 区切り）。各フォルダ直下/配下から path_hop_quality.csv を拾います。
rem 例: set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"
rem 例: スペース/日本語を含む場合 → set "FOLDERS=.\部署A\8F 東;.\部署B\10F 西"
set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"

rem タグ（; 区切り）。空なら各CSVの親フォルダ名を自動採用します。
rem 例: set "TAGS=8F-A;10F-B"
set "TAGS="

rem 1=サブフォルダも再帰、0=直下のみ
set "RECURSE=1"

rem 出力先（相対パス可）
set "OUTPUT=.\merged_teams_net_quality.csv"

rem ================ ここから処理 ================
set "CSV_LIST="
set "TAG_LIST="

rem セミコロン区切りの FOLDERS を「各要素を二重引用符付き」に展開して走査
rem 例: .\A;.\B → ".\A" ".\B"
for %%D in ("%FOLDERS:;=" "%") do (
  if not "%%~D"=="" (
    if "!RECURSE!"=="1" (
      for /r "%%~fD" %%F in (path_hop_quality.csv) do (
        call :__append "%%~fF" "%%~nxD"
      )
    ) else (
      if exist "%%~fD\path_hop_quality.csv" (
        call :__append "%%~fD\path_hop_quality.csv" "%%~nxD"
      )
    )
  )
)

rem 明示 TAGS があれば置き換え（; 区切りのまま渡す）
if not "%TAGS%"=="" set "TAG_LIST=%TAGS%"

if "%CSV_LIST%"=="" (
  echo [ERROR] path_hop_quality.csv が見つかりませんでした。>&2
  exit /b 1
)

rem PowerShell 実行（名前付き引数、; 区切りで渡す）
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -InputCsvs "%CSV_LIST%" ^
  -Tags "%TAG_LIST%" ^
  -Output "%OUTPUT%"

exit /b !ERRORLEVEL!

::__append
rem %~1 = CSV のフルパス, %~2 = 既定タグ（親フォルダ名）
set "CSV=%~1"
set "DEF_TAG=%~2"

if "%CSV_LIST%"=="" (set "CSV_LIST=%CSV%") else (set "CSV_LIST=%CSV_LIST%;%CSV%")
if "%TAG_LIST%"=="" (set "TAG_LIST=%DEF_TAG%") else (set "TAG_LIST=%TAG_LIST%;%DEF_TAG%")
exit /b 0
