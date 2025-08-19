
# Merge-TeamsNet-CSV.ps1
# 位置指定(ポジショナル)を明示的にブロックし、必ず名前付き(-InputCsvs, -Tags, -Output)で受ける

[CmdletBinding()]
param(
  # --- 位置指定ブロッカー（位置引数で来たら即エラーにする） ---
  [Parameter(Position=0)]
  [AllowNull()]
  [string]$__arg0,

  [Parameter(Position=1)]
  [AllowNull()]
  [string]$__arg1,

  # --- 本来の引数（名前付きで渡すこと） ---
  [Parameter(Mandatory=$true)]
  [Alias('Input','Files','Csvs')]
  [object]$InputCsvs,

  [Parameter(Mandatory=$true)]
  [Alias('Tag','TagsList')]
  [object]$Tags,

  [Parameter()]
  [string]$Output = ".\merged_teams_net_quality.csv",

  [switch]$Utf8Bom
)

# ========== 位置指定の明示ブロック ==========
if ($PSBoundParameters.ContainsKey('__arg0') -and $null -ne $__arg0) {
  throw "このスクリプトは『名前付き引数のみ』対応です。-InputCsvs と -Tags を必ず付けてください。（例: -InputCsvs @('a.csv','b.csv') -Tags @('8F-A','10F-B')）"
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

# 入力正規化（配列/セミコロン区切り/単体 いずれも許容）
$files = To-StringArray $InputCsvs
$tags  = To-StringArray $Tags

# フルパス化＆存在チェック
$files = $files | ForEach-Object {
  $p = $_
  if(-not (Test-Path $p)){ throw "CSV not found: $p" }
  (Resolve-Path $p).Path
}

# 個数整合
if($files.Count -ne $tags.Count){
  throw "Files($($files.Count)) と Tags($($tags.Count)) の数が一致しません。-InputCsvs と -Tags を見直してください。"
}

# すべてのヘッダーの和集合（大文字小文字無視）
$cmp = [System.StringComparer]::OrdinalIgnoreCase
$all = New-Object System.Collections.Generic.HashSet[string] $cmp

# 付加するメタ列
$meta = @('probe','machine','user','tz_offset','source_file')
foreach($m in $meta){ [void]$all.Add($m) }

$datasets = @()

for($i=0; $i -lt $files.Count; $i++){
  $path = $files[$i]
  $tag  = $tags[$i]

  # UTF-8優先、失敗時は既定(Windows: CP932)で再試行
  try {
    $rows = Import-Csv -Path $path -Encoding UTF8
  } catch {
    $rows = Import-Csv -Path $path -Encoding Default
  }

  if(-not $rows -or $rows.Count -eq 0){
    Write-Warning "Empty CSV skipped: $path"
    continue
  }

  $hdr=@{}
  foreach($n in $rows[0].PSObject.Properties.Name){
    $null = $all.Add($n)
    $hdr[$n.ToLowerInvariant()] = $n
  }

  $datasets += [pscustomobject]@{
    Path   = $path
    Tag    = $tag
    Rows   = $rows
    Header = $hdr
  }
}

if($datasets.Count -eq 0){
  throw "有効な入力CSVがありませんでした。"
}

# 和集合ヘッダーを配列化（既存列→メタ列の順）
$allHeaders = @($all.ToArray() | Where-Object { $meta -notcontains $_ }) + $meta

# 出力行の生成
$out = New-Object System.Collections.Generic.List[object]
foreach($ds in $datasets){
  $machine = $env:COMPUTERNAME
  $user    = $env:USERNAME
  $tz      = [TimeZoneInfo]::Local.BaseUtcOffset.TotalHours

  foreach($r in $ds.Rows){
    $row = [ordered]@{}
    foreach($h in $allHeaders){
      if($meta -contains $h){ continue }  # メタ列は後で付与
      if($r.PSObject.Properties.Name -contains $h){
        $row[$h] = $r.$h
      } else {
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

# 出力（Excelでの再編集を考慮して既定はUTF-8、必要ならBOM付き）
$enc = if($Utf8Bom){ 'UTF8BOM' } else { 'UTF8' }
$out | Export-Csv -NoTypeInformation -Encoding $enc -Path $Output
Write-Host "Merged -> $Output (`"$($out.Count)`" rows)"