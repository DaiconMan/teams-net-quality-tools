# Merge-TeamsNet-CSV.ps1 — 名前付き引数専用
# -InputCsvs / -Tags / -Output [/ -Utf8Bom(無視)] [/ -SplitByDate] [/ -DateColumn timestamp] [/ -DateFormat yyyyMMdd] [/ -OutputDir <dir>]

[CmdletBinding()]
param(
  # 位置指定で渡されたら明示エラー
  [Parameter(Position=0)][AllowNull()][string]$__arg0,
  [Parameter(Position=1)][AllowNull()][string]$__arg1,

  # 必ず「名前付き」で指定（文字列/配列/「;」区切り いずれも可）
  [Parameter(Mandatory=$true)][object]$InputCsvs,
  [Parameter(Mandatory=$true)][object]$Tags,

  # まとめて1本に出すときの出力先（既定）
  [Parameter()][string]$Output = ".\merged_teams_net_quality.csv",

  # 互換のため残すが本スクリプトでは未使用（常にUTF-8で出力）
  [switch]$Utf8Bom,

  # 追加: 日別でファイルを分けて出力
  [switch]$SplitByDate,
  # 日付抽出に使う列名（既定: timestamp）
  [string]$DateColumn = "timestamp",
  # 出力ファイル名に使う日付フォーマット
  [string]$DateFormat = "yyyyMMdd",
  # 分割出力の出力ディレクトリ（未指定なら Output の親フォルダ）
  [string]$OutputDir
)

# 位置指定の明示ブロック
if ($PSBoundParameters.ContainsKey('__arg0') -and $null -ne $__arg0) {
  throw "このスクリプトは『名前付き引数のみ』対応です。-InputCsvs と -Tags を必ず付けてください。"
}
if ($PSBoundParameters.ContainsKey('__arg1') -and $null -ne $__arg1) {
  throw "このスクリプトは『名前付き引数のみ』対応です。-InputCsvs と -Tags を必ず付けてください。"
}

$ErrorActionPreference = 'Stop'

function To-StringArray([object]$x){
  if($null -eq $x){ return @() }
  if($x -is [string]){
    $s=$x.Trim()
    if($s -like "*;*"){ return ($s -split ';') | ForEach-Object { $_.Trim('"',' ').Trim() } }
    else{ return @($s) }
  } elseif($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) {
    $o=@(); foreach($e in $x){ $o+=@("$e") }; return $o
  } else { return @("$x") }
}

function Get-DayKey([object]$v, [string]$fmt="yyyyMMdd"){
  $s = if($v -ne $null){ [string]$v } else { '' }
  if($s -match '^\s*(\d{4})[/-](\d{2})[/-](\d{2})'){
    return ('{0}{1}{2}' -f $Matches[1],$Matches[2],$Matches[3])
  }
  try { return ([datetime]$s).ToString($fmt) } catch { return 'Unknown' }
}

# 入力正規化
$files = To-StringArray $InputCsvs
$tags  = To-StringArray $Tags

# フルパス化 & 存在検証
$files = foreach($p in $files){
  if(-not (Test-Path -LiteralPath $p)){ throw "CSV not found: $p" }
  (Resolve-Path -LiteralPath $p).Path
}
if($files.Count -ne $tags.Count){
  throw "Files($($files.Count)) と Tags($($tags.Count)) の数が一致しません。-InputCsvs と -Tags を見直してください。"
}

# ヘッダー和集合（大文字小文字無視）
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$all = New-Object System.Collections.Generic.HashSet[string] $cmp

# 付与するメタ列（※ machine / user は含めない）
$meta = @('probe','tz_offset','source_file')
foreach($m in $meta){ [void]$all.Add($m) }

$datasets = @()

for($i=0;$i -lt $files.Count;$i++){
  $path = $files[$i]; $tag  = $tags[$i]

  try   { $rows = Import-Csv -Path $path -Encoding UTF8 }
  catch { $rows = Import-Csv -Path $path -Encoding Default }

  if(-not $rows -or $rows.Count -eq 0){
    Write-Warning "Empty CSV skipped: $path"
    continue
  }

  foreach($n in $rows[0].PSObject.Properties.Name){
    $null = $all.Add($n)
  }

  $datasets += [pscustomobject]@{
    Path = $path
    Tag  = $tag
    Rows = $rows
  }
}
if($datasets.Count -eq 0){ throw "有効な入力CSVがありませんでした。" }

# 和集合ヘッダー（既存列→メタ列）
$existingHeaders = @()
foreach($h in $all){ if($meta -notcontains $h){ $existingHeaders += $h } }
$allHeaders = @($existingHeaders + $meta)

# 出力生成（SplitByDate の場合は日単位で分配）
if($SplitByDate){
  # 出力ディレクトリを決定
  $baseDir = if($OutputDir){
    $OutputDir
  } else {
    $parent = Split-Path -Parent $Output
    if([string]::IsNullOrWhiteSpace($parent)){ '.' } else { $parent }
  }
  if(-not (Test-Path -LiteralPath $baseDir)){ New-Item -ItemType Directory -Path $baseDir | Out-Null }

  # dayKey => List[pscustomobject]
  $groups = @{}

  foreach($ds in $datasets){
    $tz = [TimeZoneInfo]::Local.BaseUtcOffset.TotalHours
    foreach($r in $ds.Rows){
      $row = [ordered]@{}
      foreach($h in $allHeaders){
        if($meta -contains $h){ continue }
        if($r.PSObject.Properties.Name -contains $h){
          $row[$h] = $r.$h
        } else {
          $k = $r.PSObject.Properties.Name | Where-Object { $_ -ieq $h } | Select-Object -First 1
          $row[$h] = if($k){ $r.$k } else { $null }
        }
      }
      # メタ列（machine/user は出さない）
      $row['probe']       = $ds.Tag
      $row['tz_offset']   = $tz
      $row['source_file'] = $ds.Path

      $obj = [pscustomobject]$row
      $key = if($row.Contains($DateColumn)){ Get-DayKey $row[$DateColumn] $DateFormat } else { 'Unknown' }

      if(-not $groups.ContainsKey($key)){
        $groups[$key] = New-Object System.Collections.Generic.List[object]
      }
      [void]$groups[$key].Add($obj)
    }
  }

  # 書き出し（キーごとにファイル）
  foreach($k in ($groups.Keys | Sort-Object)){
    $file = Join-Path $baseDir ("merged_{0}.csv" -f $k)
    $groups[$k] | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $file
    Write-Host "Merged -> $file (`"$($groups[$k].Count)`" rows)"
  }
}
else{
  # 1本にまとめて出力
  $out = New-Object System.Collections.Generic.List[object]
  foreach($ds in $datasets){
    $tz = [TimeZoneInfo]::Local.BaseUtcOffset.TotalHours
    foreach($r in $ds.Rows){
      $row = [ordered]@{}
      foreach($h in $allHeaders){
        if($meta -contains $h){ continue }
        if($r.PSObject.Properties.Name -contains $h){
          $row[$h] = $r.$h
        } else {
          $k = $r.PSObject.Properties.Name | Where-Object { $_ -ieq $h } | Select-Object -First 1
          $row[$h] = if($k){ $r.$k } else { $null }
        }
      }
      $row['probe']       = $ds.Tag
      $row['tz_offset']   = $tz
      $row['source_file'] = $ds.Path
      [void]$out.Add([pscustomobject]$row)
    }
  }
  $out | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $Output
  Write-Host "Merged -> $Output (`"$($out.Count)`" rows)"
}
