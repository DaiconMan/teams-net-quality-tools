param(
  # 位置0: CSVパス（1本 or 複数）。";" 区切りも可
  [Parameter(Mandatory=$true, Position=0)]
  [Alias('Input','Files','Csvs')]
  [object]$InputCsvs,

  # 位置1: タグ（1つ or 複数）。";" 区切りも可
  [Parameter(Mandatory=$true, Position=1)]
  [Alias('Tag','TagsList')]
  [object]$Tags,

  # 位置2: 出力ファイル
  [Parameter(Position=2)]
  [string]$Output = ".\merged_teams_net_quality.csv",

  [switch]$Utf8Bom
)

function To-StringArray([object]$x){
  if($null -eq $x){ return @() }
  if($x -is [string]){
    $s=$x.Trim()
    if($s -like "*;*"){ return ($s -split ';') | ForEach-Object { $_.Trim('"',' ').Trim() } }
    else{ return @($s) }
  }
  elseif($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])){
    $o=@(); foreach($e in $x){ $o+=@("$e") }; return $o
  }
  else{ return @("$x") }
}

# 正規化
$files = To-StringArray $InputCsvs
$tags  = To-StringArray $Tags

# フルパス化＆存在チェック
$files = $files | ForEach-Object {
  $p = $_
  if(-not (Test-Path $p)){ throw "CSV not found: $p" }
  (Resolve-Path $p).Path
}

if($files.Count -ne $tags.Count){
  throw "Files($($files.Count)) と Tags($($tags.Count)) の数が一致しません。"
}

$ErrorActionPreference = 'Stop'

# すべてのヘッダーの和集合を作る（大文字小文字は区別しない）
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$all = New-Object System.Collections.Generic.HashSet[string] $cmp
# 付加するメタ列
$meta = @('probe','machine','user','tz_offset','source_file')
foreach($m in $meta){ [void]$all.Add($m) }

$datasets = @()

for($i=0; $i -lt $InputCsvs.Count; $i++){
  $path = (Resolve-Path $InputCsvs[$i]).Path
  if(-not (Test-Path $path)){ throw "CSV not found: $path" }

  # UTF-8で読む（Shift-JIS の場合は先に変換推奨）
  $rows = Import-Csv -Path $path -Encoding UTF8
  if(-not $rows){ continue }

  # 見つけたヘッダーを集合に追加
  $hdr = @{}
  foreach($n in $rows[0].PSObject.Properties.Name){
    $null = $all.Add($n)
    $hdr[$n.ToLowerInvariant()] = $n
  }

  $datasets += [pscustomobject]@{
    Path   = $path
    Tag    = $Tags[$i]
    Rows   = $rows
    Header = $hdr
  }
}

# 和集合ヘッダーを配列化（順番：既存列→メタ列の順）
$allHeaders = @($all.ToArray() | Where-Object { $meta -notcontains $_ }) + $meta

# 出力用に行を正規化
$out = New-Object System.Collections.Generic.List[object]
foreach($ds in $datasets){
  $machine = $env:COMPUTERNAME
  $user    = $env:USERNAME
  $tz      = [TimeZoneInfo]::Local.BaseUtcOffset.TotalHours

  foreach($r in $ds.Rows){
    $row = [ordered]@{}
    foreach($h in $allHeaders){
      if($meta -contains $h){ continue }  # メタ列は後で埋める
      if($r.PSObject.Properties.Name -contains $h){
        $row[$h] = $r.$h
      } else {
        # 大文字小文字違いなどにも一応対応
        $k = $r.PSObject.Properties.Name | Where-Object { $_ -ieq $h } | Select-Object -First 1
        $row[$h] = if($k){ $r.$k } else { $null }
      }
    }
    # メタ列
    $row['probe']       = $ds.Tag
    $row['machine']     = $machine
    $row['user']        = $user
    $row['tz_offset']   = $tz
    $row['source_file'] = $ds.Path

    $out.Add([pscustomobject]$row) | Out-Null
  }
}

# 書き出し（UTF-8 / 必要ならBOM）
$enc = if($Utf8Bom){ 'UTF8BOM' } else { 'UTF8' }
$out | Export-Csv -NoTypeInformation -Encoding $enc -Path $Output
Write-Host "Merged -> $Output (`"$($out.Count)`" rows)"