# Merge-TeamsNet-CSV.ps1
# 目的:
#   - 入力で受け取った「パス」は“ファイル or ディレクトリ”のどちらでも可
#   - ディレクトリが渡された場合は、その配下の *.csv（既定は非再帰）を収集してマージ
#   - 値の加工は行わない（単純マージ）。列構成は現状維持（dest_ip を含み、無い場合は空欄）
#   - SplitByDate 機能は従来通りサポート
# 注意:
#   - PowerShell 5.1 対応（?: 不使用）、$Host 予約語は未使用
#   - OneDrive / 日本語 / スペースを含むパスに配慮（-LiteralPath / Resolve-Path）
#   - Cドライブ非表示環境を想定（相対/OneDriveパスで動作）

[CmdletBinding()]
param(
  # ; 区切り or 配列で、CSVファイル or ディレクトリを列挙
  [Parameter(Mandatory=$true)][object]$InputCsvs,
  # 入力単位（InputCsvs の各要素）に対応するタグ。空なら自動（末端フォルダ名）
  [Parameter(Mandatory=$true)][object]$Tags,

  # 単一ファイル出力 or 日別分割
  [Parameter()][string]$Output = ".\merged_teams_net_quality.csv",
  [Parameter()][switch]$SplitByDate,
  [Parameter()][string]$DateColumn = "timestamp",
  [Parameter()][string]$DateFormat = "yyyyMMdd",
  [Parameter()][string]$OutputDir,

  # ディレクトリ指定時に再帰的に *.csv を拾うか
  [Parameter()][switch]$Recurse
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
  if($s -match '^\s*(\d{4})[/-](\d{2})[/-](\d{2})'){ return ('{0}{1}{2}' -f $Matches[1],$Matches[2],$Matches[3]) }
  try { return ([datetime]$s).ToString($fmt) } catch { return "Unknown" }
}

# 拡張子 ".cs　v" のような誤記を ".csv" に矯正（末尾のみ）
function Fix-CsvExtension([string]$p){
  if([string]::IsNullOrWhiteSpace($p)){ return $p }
  # 全角/半角スペースを含む ".cs   v" / ".cs　　v" を ".csv" に置換
  $p2 = [regex]::Replace($p, '\.cs[\u3000\s]+v$', '.csv')
  return $p2
}

# 入力単位（ファイル or ディレクトリ）を解析し、実ファイル一覧と単位タグの対応を展開
function Expand-InputUnits([string[]]$units, [string[]]$unitTags, [switch]$recurse){
  $out = New-Object System.Collections.Generic.List[object]
  for($i=0;$i -lt $units.Count;$i++){
    $raw = $units[$i]
    # 末尾の ".cs [spaces] v" を補正
    $maybe = Fix-CsvExtension -p $raw

    # 前後空白は削除（全角/半角）
    $trimmed = $maybe.Trim((" `t`r`n") + ([char]0x3000))
    $path = $trimmed
    $tag  = $unitTags[$i]

    # ファイル/ディレクトリの存在判定（存在しない場合は、そのまま警告）
    $isFile = $false; $isDir = $false
    if(Test-Path -LiteralPath $path){
      $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
      if($item -ne $null){
        if($item.PSIsContainer){ $isDir = $true } else { $isFile = $true }
      }
    }

    if($isFile){
      $rp = (Resolve-Path -LiteralPath $path).Path
      $out.Add([pscustomobject]@{ File=$rp; Tag=$tag }) | Out-Null
    } elseif($isDir){
      $opt = @{}
      $opt['LiteralPath'] = $path
      $opt['File'] = $true
      $opt['Filter'] = '*.csv'
      if($recurse){ $opt['Recurse'] = $true }
      $files = Get-ChildItem @opt | Select-Object -ExpandProperty FullName
      if(-not $files -or $files.Count -eq 0){
        Write-Warning ("CSV not found in directory: {0}" -f $path)
      } else {
        foreach($f in $files){
          $out.Add([pscustomobject]@{ File=$f; Tag=$tag }) | Out-Null
        }
      }
    } else {
      Write-Warning ("Input path not found (skipped): {0}" -f $path)
    }
  }
  return ,$out
}

# -------- main --------
$units = To-StringArray $InputCsvs
$unitTags = To-StringArray $Tags

# タグ補完（空なら入力単位の末端フォルダ名）
if($unitTags.Count -eq 0){
  $unitTags = @()
  foreach($u in $units){
    $p = $u
    $p = Fix-CsvExtension -p $p
    $p = $p.Trim((" `t`r`n") + ([char]0x3000))
    if(Test-Path -LiteralPath $p){
      $gi = Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
      if($gi -and -not $gi.PSIsContainer){
        $base = Split-Path -Parent $gi.FullName
        $unitTags += (Split-Path -Leaf $base)
      } else {
        $unitTags += (Split-Path -Leaf $p)
      }
    } else {
      # 見つからない場合は文字列から推定
      $unitTags += (Split-Path -Leaf $p)
    }
  }
}
if($units.Count -ne $unitTags.Count){
  throw "InputCsvs の要素数($($units.Count)) と Tags の要素数($($unitTags.Count)) が一致しません。"
}

# 単位を展開（ディレクトリ→*.csv 群）
$entries = Expand-InputUnits -units $units -unitTags $unitTags -recurse:$Recurse
if($entries.Count -eq 0){ throw "有効な CSV が見つかりませんでした。" }

# 列集合を構築（大小無視、現状維持：dest_ip を強制含有、メタ列は末尾）
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$headerSet = New-Object System.Collections.Generic.HashSet[string] $cmp
$forceColumns = @('dest_ip')                   # 列は残す。無ければ空欄で出力
$metaColumns  = @('probe','tz_offset','source_file')

# 最初のファイルの列順を優先して安定化
$firstFile = $entries[0].File
try   { $firstRows = Import-Csv -Path $firstFile -Encoding UTF8 }
catch { $firstRows = Import-Csv -Path $firstFile -Encoding Default }
if($firstRows -and $firstRows.Count -gt 0){
  foreach($n in $firstRows[0].PSObject.Properties.Name){ [void]$headerSet.Add($n) }
}

# 全ファイルからヘッダ和集合
foreach($e in $entries){
  try   { $rows = Import-Csv -Path $e.File -Encoding UTF8; }
  catch { $rows = Import-Csv -Path $e.File -Encoding Default; }
  if($rows -and $rows.Count -gt 0){
    foreach($n in $rows[0].PSObject.Properties.Name){ [void]$headerSet.Add($n) }
  }
}
foreach($fc in $forceColumns){ [void]$headerSet.Add($fc) }

# 最終ヘッダ順を決定
$ordered = @()
if($firstRows -and $firstRows.Count -gt 0){
  foreach($n in $firstRows[0].PSObject.Properties.Name){
    if($headerSet.Contains($n) -and -not $ordered.Contains($n)){ $ordered += $n }
  }
}
foreach($h in $headerSet){ if(-not ($ordered -contains $h)){ $ordered += $h } }
$allHeaders = @($ordered + $metaColumns)

# 本体マージ
$outList = New-Object System.Collections.Generic.List[object]
$tz = ([datetimeoffset](Get-Date)).ToString("zzz")

foreach($e in $entries){
  try   { $rows = Import-Csv -Path $e.File -Encoding UTF8 }
  catch { $rows = Import-Csv -Path $e.File -Encoding Default }

  if(-not $rows -or $rows.Count -eq 0){
    Write-Warning ("Empty CSV skipped: {0}" -f $e.File)
    continue
  }

  foreach($r in $rows){
    $row = [ordered]@{}
    foreach($h in $ordered){
      if($r.PSObject.Properties.Name -contains $h){
        $row[$h] = $r.$h
      } else {
        $k = $r.PSObject.Properties.Name | Where-Object { $_ -ieq $h } | Select-Object -First 1
        if($k){ $row[$h] = $r.$k } else { $row[$h] = $null }
      }
    }
    $row['probe']       = $e.Tag
    $row['tz_offset']   = $tz
    $row['source_file'] = $e.File

    [void]$outList.Add([pscustomobject]$row)
  }
}

if($outList.Count -eq 0){ throw "有効な行がありませんでした。" }

# 出力
if($SplitByDate){
  $baseDir = $OutputDir
  if([string]::IsNullOrWhiteSpace($baseDir)){
    $p = Split-Path -Parent $Output
    if([string]::IsNullOrWhiteSpace($p)){ $baseDir = '.' } else { $baseDir = $p }
  }
  if(-not (Test-Path -LiteralPath $baseDir)){ New-Item -ItemType Directory -Path $baseDir | Out-Null }

  foreach($o in $outList){
    $day = "Unknown"
    $dc = $DateColumn
    if($o.PSObject.Properties.Name -contains $dc){ $day = Get-DayKey $o.$dc $DateFormat }
    $file = Join-Path $baseDir ("merged_{0}.csv" -f $day)
    $exists = Test-Path -LiteralPath $file
    $o | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8 -Append:($exists) -Force
  }
  Write-Host ("Split merged -> {0}" -f $baseDir)
} else {
  $outList | Export-Csv -Path $Output -NoTypeInformation -Encoding UTF8
  Write-Host ("Merged -> {0} (""{1}"" rows)" -f $Output, $outList.Count)
}
