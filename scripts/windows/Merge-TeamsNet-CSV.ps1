# Merge-TeamsNet-CSV.ps1
# 目的: 各入力フォルダ内の path_hop_quality.csv（等、指定CSV）を**加工せず**マージする。
# 注意: カラムは現状維持（dest_ip を含む）。入力に無ければ空欄で出力。
# PS 5.1 対応（?: 不使用）、$Host 変数は未使用。OneDrive/日本語パス対応。

[CmdletBinding()]
param(
  # ; 区切り or 配列 で CSV ファイルパス（例: .\A\path_hop_quality.csv;.\B\path_hop_quality.csv）
  [Parameter(Mandatory=$true)][object]$InputCsvs,
  # 各CSVに対応するタグ（; 区切り or 配列）。空なら親フォルダ名を自動採用
  [Parameter(Mandatory=$true)][object]$Tags,

  # 1本出力 or 日別分割
  [Parameter()][string]$Output = ".\merged_teams_net_quality.csv",
  [Parameter()][switch]$SplitByDate,
  [Parameter()][string]$DateColumn = "timestamp",
  [Parameter()][string]$DateFormat = "yyyyMMdd",
  [Parameter()][string]$OutputDir
)

# -------- helpers --------
function To-StringArray([object]$v){
  if($null -eq $v){ return @() }
  if($v -is [System.Array]){ return @($v | ForEach-Object { [string]$_ }) }
  $s = [string]$v
  if([string]::IsNullOrWhiteSpace($s)){ return @() }
  if($s.Contains(";")){ return @($s.Split(";",[System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }) }
  return @($s)
}

function Get-DayKey([object]$v, [string]$fmt="yyyyMMdd"){
  $s = if($v -ne $null){ [string]$v } else { "" }
  if($s -match '^\s*(\d{4})[/-](\d{2})[/-](\d{2})'){
    return ('{0}{1}{2}' -f $Matches[1],$Matches[2],$Matches[3])
  }
  try { return ([datetime]$s).ToString($fmt) } catch { return "Unknown" }
}

# -------- main --------
$files = To-StringArray $InputCsvs
$tags  = To-StringArray $Tags

# 検証 & 正規化（フルパス化）
$resolved = New-Object System.Collections.Generic.List[string]
foreach($p in $files){
  if(-not (Test-Path -LiteralPath $p)){ throw "CSV not found: $p" }
  $rp = (Resolve-Path -LiteralPath $p).Path
  [void]$resolved.Add($rp)
}
$files = @($resolved)

# タグ補完（空なら末端フォルダ名）
if($tags.Count -eq 0){
  $tags = @()
  foreach($f in $files){
    $dir = Split-Path -Parent $f
    $tags += (Split-Path -Leaf $dir)
  }
}
if($files.Count -ne $tags.Count){
  throw "Files($($files.Count)) と Tags($($tags.Count)) の数が一致しません。-InputCsvs と -Tags を見直してください。"
}

# ヘッダー和集合（大小無視）。現状維持カラムを**強制**含める。
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$headerSet = New-Object System.Collections.Generic.HashSet[string] $cmp

# 現状維持のために必ず残す列（dest_ip とメタ列）
$forceColumns = @('dest_ip') # ※ 入力に無い場合は空で出力（列構成維持）
$metaColumns  = @('probe','tz_offset','source_file') # 既存どおり末尾に付与

foreach($f in $files){
  try   { $rows = Import-Csv -Path $f -Encoding UTF8 }
  catch { $rows = Import-Csv -Path $f -Encoding Default }
  if(-not $rows -or $rows.Count -eq 0){ continue }
  foreach($n in $rows[0].PSObject.Properties.Name){ [void]$headerSet.Add($n) }
}
foreach($fc in $forceColumns){ [void]$headerSet.Add($fc) }

# 最終ヘッダー（入力順に近い安定化：最初のCSVの順 → 和集合残り → 強制列 → メタ列）
$ordered = @()
if($files.Count -gt 0){
  try   { $firstRows = Import-Csv -Path $files[0] -Encoding UTF8 }
  catch { $firstRows = Import-Csv -Path $files[0] -Encoding Default }
  if($firstRows -and $firstRows.Count -gt 0){
    foreach($n in $firstRows[0].PSObject.Properties.Name){
      if($headerSet.Contains($n) -and -not $ordered.Contains($n)){ $ordered += $n }
    }
  }
}
# 和集合の残り
foreach($h in $headerSet){
  if(-not ($ordered -contains $h)){ $ordered += $h }
}

# メタ列は最後
$allHeaders = @($ordered + $metaColumns)

# 出力準備
$outList = New-Object System.Collections.Generic.List[object]
$tz = ([datetimeoffset](Get-Date)).ToString("zzz")

for($i=0; $i -lt $files.Count; $i++){
  $path = $files[$i]; $tag = $tags[$i]
  try   { $rows = Import-Csv -Path $path -Encoding UTF8 }
  catch { $rows = Import-Csv -Path $path -Encoding Default }

  if(-not $rows -or $rows.Count -eq 0){
    Write-Warning "Empty CSV skipped: $path"
    continue
  }

  foreach($r in $rows){
    $row = [ordered]@{}
    foreach($h in $ordered){
      if($r.PSObject.Properties.Name -contains $h){
        $row[$h] = $r.$h
      } else {
        # 大文字小文字違いを吸収
        $k = $r.PSObject.Properties.Name | Where-Object { $_ -ieq $h } | Select-Object -First 1
        if($k){ $row[$h] = $r.$k } else { $row[$h] = $null }
      }
    }
    # 現状どおりメタ列を最後に付与（machine/user は出力しない仕様）
    $row['probe']       = $tag
    $row['tz_offset']   = $tz
    $row['source_file'] = $path

    [void]$outList.Add([pscustomobject]$row)
  }
}

if($outList.Count -eq 0){ throw "有効な入力CSVがありませんでした。" }

if($SplitByDate){
  $baseDir = $OutputDir
  if([string]::IsNullOrWhiteSpace($baseDir)){
    $p = Split-Path -Parent $Output
    if([string]::IsNullOrWhiteSpace($p)){ $baseDir = '.' } else { $baseDir = $p }
  }
  if(-not (Test-Path -LiteralPath $baseDir)){ New-Item -ItemType Directory -Path $baseDir | Out-Null }

  # ヘッダーは Export-Csv に任せる（追記）
  foreach($o in $outList){
    $day = "Unknown"
    $dc = $DateColumn
    if($o.PSObject.Properties.Name -contains $dc){
      $day = Get-DayKey $o.$dc $DateFormat
    }
    $file = Join-Path $baseDir ("merged_{0}.csv" -f $day)
    $exists = Test-Path -LiteralPath $file
    $o | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 -Append:($exists) -Force
  }
  Write-Host ("Split merged -> {0}" -f $baseDir)
} else {
  $outList | Export-Csv -Path $Output -NoTypeInformation -Encoding UTF8
  Write-Host ("Merged -> {0} (""{1}"" rows)" -f $Output, $outList.Count)
}
