@echo off
setlocal EnableExtensions
chcp 932 >nul 2>&1

rem ===== 設定 =====
set "SCRIPT=Merge-TeamsNet-CSV.ps1"

rem ; 区切りでフォルダを列挙（相対パスOK、内側に追加の " は不要）
rem 例: set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"
set "FOLDERS=.\Logs\8F-A;.\Logs\10F-B"

rem 各フォルダのタグ。空なら末端フォルダ名を自動採用（8F-A など）
set "TAGS="

rem 1=サブフォルダ再帰 / 0=直下のみ
set "RECURSE=1"

rem 収集するファイルパターン
set "PATTERN=*.csv"

rem 単一出力のときの書き出し先（相対ならこのbat基準）
set "OUTPUT=merged_all.csv"

rem --- 日別分割オプション ---
set "SPLIT=1"                 rem 1=日別に分割、0=分割しない
set "DATECOL=timestamp"       rem 日付抽出に使う列名
set "DATEFMT=yyyyMMdd"        rem 出力ファイル名の日付フォーマット
set "OUTDIR=merged_by_day"    rem 分割出力のフォルダ（相対/絶対どちらでも可）
rem ===========================

set "BASE=%~dp0"
pushd "%BASE%" >nul 2>&1

rem PowerShell 実体
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

echo --- 収集とマージを開始 ---
"%PS%" -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$mergeScript=[IO.Path]::GetFullPath('%ABS_SCRIPT%');" ^
  "$out=[IO.Path]::GetFullPath('%ABS_OUT%');" ^
  "$raw = $env:FOLDERS -split ';';" ^
  "$folders = foreach($x in $raw){ $t=$x.Trim(); if($t){ if(Test-Path -LiteralPath $t){ (Resolve-Path -LiteralPath $t).Path } else { Write-Warning ('Skip (not found): {0}' -f $x) } } };" ^
  "if(-not $folders -or $folders.Count -eq 0){ throw 'FOLDERS で指定されたフォルダが見つかりません。' }" ^
  "$tags = if([string]::IsNullOrWhiteSpace($env:TAGS)){ $folders | ForEach-Object { Split-Path $_ -Leaf } } else { ($env:TAGS -split ';' | ForEach-Object { $_.Trim() }) };" ^
  "if($folders.Count -ne $tags.Count){ throw 'FOLDERS と TAGS の数が一致しません。TAGS を空にするとフォルダ名が自動タグになります。' }" ^
  "$all    = New-Object 'System.Collections.Generic.List[string]';" ^
  "$tagsEx = New-Object 'System.Collections.Generic.List[string]';" ^
  "for($i=0;$i -lt $folders.Count;$i++){" ^
  "  $f=$folders[$i]; $tag=$tags[$i];" ^
  "  $opt=@{LiteralPath=$f;Filter='%PATTERN%';File=$true}; if('%RECURSE%' -eq '1'){ $opt.Recurse=$true }" ^
  "  $files = Get-ChildItem @opt | Select-Object -Expand FullName;" ^
  "  foreach($p in $files){ [void]$all.Add($p); [void]$tagsEx.Add($tag) }" ^
  "}" ^
  "if($all.Count -eq 0){ throw '指定フォルダ群に CSV が見つかりません。' }" ^
  "$inputs = [string]::Join(';', @($all));" ^
  "$tagsStr= [string]::Join(';', @($tagsEx));" ^
  "$params = @{ InputCsvs = $inputs; Tags = $tagsStr; Output = $out };" ^
  "if('%SPLIT%' -eq '1'){" ^
  "  $params['SplitByDate'] = $true;" ^
  "  if(-not [string]::IsNullOrWhiteSpace('%DATECOL%')){  $params['DateColumn']  = '%DATECOL%' }" ^
  "  if(-not [string]::IsNullOrWhiteSpace('%DATEFMT%')){  $params['DateFormat']  = '%DATEFMT%' }" ^
  "  if(-not [string]::IsNullOrWhiteSpace('%OUTDIR%')){   $params['OutputDir']   = [IO.Path]::GetFullPath('%OUTDIR%') }" ^
  "}" ^
  "& (Get-Item -LiteralPath $mergeScript).FullName @params" ^
  "; if($?) { if('%SPLIT%' -eq '1'){ Write-Host ('* 分割出力完了 -> {0}' -f ($params.OutputDir ?? Split-Path -Parent $params.Output)) } else { Write-Host ('* 出力: {0}' -f $params.Output) }; exit 0 } else { exit 1 }"

set "RC=%ERRORLEVEL%"
echo * 終了コード: %RC%
if not "%RC%"=="0" echo [ERROR] マージに失敗しました。上のメッセージを確認してください.
goto :end

:fail
set "RC=1"

:end
popd >nul 2>&1
exit /b %RC%
