@echo off
setlocal EnableExtensions

rem ===== 設定 =====
set "SCRIPT=Merge-TeamsNet-CSV.ps1"
set "FOLDERS=C:\Logs\8F-A;C:\Logs\10F-B"  rem ; 区切りで複数
set "TAGS="                                 rem 空なら各フォルダ名をタグ化
set "RECURSE=1"                             rem 1=再帰, 0=直下のみ
set "PATTERN=*.csv"
set "OUTPUT=merged_all.csv"
set "UTF8BOM=1"
rem =================

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
  "$all    = New-Object 'System.Collections.Generic.List[string]';" ^
  "$tagsEx = New-Object 'System.Collections.Generic.List[string]';" ^
  "for($i=0;$i -lt $folders.Count;$i++){" ^
  "  $f=$folders[$i]; $tag=$tags[$i];" ^
  "  $opt=@{Path=$f;Filter='%PATTERN%';File=$true}; if('%RECURSE%' -eq '1'){ $opt.Recurse=$true }" ^
  "  $files = Get-ChildItem @opt | Select-Object -Expand FullName;" ^
  "  foreach($p in $files){ [void]$all.Add($p); [void]$tagsEx.Add($tag) }" ^
  "}" ^
  "if($all.Count -eq 0){ throw '指定フォルダ群に CSV が見つかりません。' }" ^
  "$inputs = [string]::Join(';', @($all));" ^
  "$tagsStr= [string]::Join(';', @($tagsEx));" ^
  "Write-Host ('* 収集ファイル数: {0} / タグ数: {1}' -f $all.Count, $tagsEx.Count);" ^
  "if('%UTF8BOM%' -eq '1'){" ^
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