# Merge-TeamsNet-CSV.ps1  --  名前付き引数専用: -InputCsvs / -Tags / -Output [/ -Utf8Bom]
[CmdletBinding()]
param(
  # 位置指定で渡されたら明示エラー
  [Parameter(Position=0)][AllowNull()][string]$__arg0,
  [Parameter(Position=1)][AllowNull()][string]$__arg1,

  # ← 必ず「名前付き」で指定（文字列/配列/「;」区切り いずれも可）
  [Parameter(Mandatory=$true)]
  [object]$InputCsvs,
  [Parameter(Mandatory=$true)]
  [object]$Tags,

  [Parameter()]
  [string]$Output = ".\merged_teams_net_quality.csv",

  [switch]$Utf8Bom
)

# 位置指定ブロック
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

# 入力正規化
$files = To-StringArray $InputCsvs
$tags  = To-StringArray $Tags

# フルパス化 & 存在検証
$files = $files | ForEach-Object {
  if(-not (Test-Path $_)){ throw "CSV not found: $_" }
  (Resolve-Path $_).Path
}

if($files.Count -ne $tags.Count){
  throw "Files($($files.Count)) と Tags($($tags.Count)) の数が一致しません。-InputCsvs と -Tags を見直してください。"
}

# ヘッダー和集合（大文字小文字無視）
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$all = New-Object System.Collections.Generic.HashSet[string] $cmp

# 付与するメタ列
$meta = @('probe','machine','user','tz_offset','source_file')
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
$allHeaders = @($all.ToArray() | Where-Object { $meta -notcontains $_ }) + $meta

# 出力行生成
$out = New-Object System.Collections.Generic.List[object]
foreach($ds in $datasets){
  $machine = $env:COMPUTERNAME
  $user    = $env:USERNAME
  $tz      = [TimeZoneInfo]::Local.BaseUtcOffset.TotalHours

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
    $row['machine']     = $machine
    $row['user']        = $user
    $row['tz_offset']   = $tz
    $row['source_file'] = $ds.Path

    $out.Add([pscustomobject]$row) | Out-Null
  }
}

# 書き出し
$enc = if($Utf8Bom){ 'UTF8BOM' } else { 'UTF8' }
$out | Export-Csv -NoTypeInformation -Encoding $enc -Path $Output
Write-Host "Merged -> $Output (`"$($out.Count)`" rows)"