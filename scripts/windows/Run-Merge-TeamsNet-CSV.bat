@echo off
setlocal EnableExtensions

rem ===== 設定（必要に応じて変更）=================================
set "SCRIPT=Merge-TeamsNet-CSV.ps1"

rem 1つ以上のフォルダを ; で区切って指定（例：8F-A と 10F-B）
set "FOLDERS=C:\Logs\8F-A;C:\Logs\10F-B"

rem 各フォルダのタグ。空なら「フォルダ名」を自動タグ化（例：8F-A / 10F-B）
set "TAGS="

rem 再帰的に集めるなら 1、直下のみなら 0
set "RECURSE=1"

rem 収集するパターン
set "PATTERN=*.csv"

rem 出力（相対ならこのbatのあるフォルダ基準）
set "OUTPUT=merged_all.csv"

rem Excel互換のためBOM付きUTF-8で出力するなら 1
set "UTF8BOM=1"
rem ===============================================================

chcp 932 >nul 2>&1

set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1

set "PS=pwsh.exe"
where pwsh.exe >nul 2>&1 || set "PS=powershell.exe"

for %%I in ("%SCRIPT%")  do set "ABS_SCRIPT=%%~fI"
for %%I in ("%OUTPUT%")  do set "ABS_OUT=%%~fI"

if not exist "%ABS_SCRIPT%" (
  echo [ERROR] スクリプトが見つかりません: "%ABS_SCRIPT%"
  goto :fail
)
if not defined FOLDERS (
  echo [ERROR] FOLDERS が未設定です。; 区切りでフォルダを指定してください。
  goto :fail
)

set "BOMSW="
if "%UTF8BOM%"=="1" set "BOMSW=-Utf8Bom"

echo --- 収集とマージを開始 ---
"%PS%" -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$script=[IO.Path]::GetFullPath('%ABS_SCRIPT%');" ^
  "$out=[IO.Path]::GetFullPath('%ABS_OUT%');" ^
  "$folders=($env:FOLDERS -split ';' | Where-Object { $_ -and (Test-Path $_) });" ^
  "if($folders.Count -eq 0){ throw 'FOLDERS で指定されたフォルダが見つかりません。' }" ^
  "$tags = if([string]::IsNullOrWhiteSpace($env:TAGS)){ $folders | ForEach-Object { Split-Path $_ -Leaf } } else { $env:TAGS -split ';' };" ^
  "if($folders.Count -ne $tags.Count){ throw 'FOLDERS と TAGS の数が一致しません。TAGS を空にするとフォルダ名が自動タグになります。' }" ^
  "$all = New-Object 'System.Collections.Generic.List[string]';" ^
  "$tagsExp = New-Object 'System.Collections.Generic.List[string]';" ^
  "for($i=0;$i -lt $folders.Count;$i++){" ^
  "  $f=$folders[$i]; $tag=$tags[$i];" ^
  "  $opt=@{Path=$f;Filter=$env:PATTERN;File=$true}; if($env:RECURSE -eq '1'){ $opt.Recurse=$true }" ^
  "  $files = Get-ChildItem @opt | Select-Object -Expand FullName;" ^
  "  foreach($p in $files){ $all.Add($p); $tagsExp.Add($tag) }" ^
  "}" ^
  "if($all.Count -eq 0){ throw '指定フォルダ群に CSV が見つかりません。' }" ^
  "$inputs = ($all.ToArray() -join ';'); $tagsStr = ($tagsExp.ToArray() -join ';');" ^
  "Write-Host ('* 収集ファイル数: {0} / タグ数: {1}' -f $all.Count, $tagsExp.Count);" ^
  "if($env:UTF8BOM -eq '1'){" ^
  "  & $script -InputCsvs $inputs -Tags $tagsStr -Output $out -Utf8Bom" ^
  "} else {" ^
  "  & $script -InputCsvs $inputs -Tags $tagsStr -Output $out" ^
  "}" ^
  "; if($?) { Write-Host ('* 出力: {0}' -f $out); exit 0 } else { exit 1 }"

set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] マージに失敗しました。上のメッセージを確認してください.
goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
exit /b %RC%