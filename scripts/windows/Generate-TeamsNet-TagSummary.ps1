<#
.SYNOPSIS
  Tag(=測定場所)ごとの品質サマリを作成し、Excel/CSVで出力する（PowerShell 5.1対応）

.DESCRIPTION
  Run-Merge-TeamsNet-CSV.bat の出力（例: rtt_ms_gateway, rtt_ms_hop2 などの横持ち）にも整合。
  レイテンシ候補列を自動検出し、必要に応じてロング化（列名から Target を推定）して集計。
  出力先（-Output）が未指定または相対パスの場合は、スクリプトの2つ上（リポジトリ直下）を基準に解決。
  多段フォルダも [System.IO.Directory]::CreateDirectory() で必ず作成。

.PARAMETER CsvPath
  マージ済み入力CSV（Run-Merge-TeamsNet-CSV.bat の出力を想定）。相対ならリポジトリ直下基準。

.PARAMETER Output
  出力フォルダ（既定: <repo>\Output）。相対ならリポジトリ直下基準。

.PARAMETER TagColumn
  Tag 列の明示（既定: 自動検出: probe|tag|location|site|place）

.PARAMETER TargetColumn
  1列レイテンシ時の Target 列（既定: 自動検出: target|hostname|endpoint|server|dst|fqdn|ip|addr）

.PARAMETER HostColumn
  互換のため受け付け（内部では Target として扱う。TargetColumn が優先）

.PARAMETER LatencyColumn
  レイテンシ列の明示（既定: 自動検出。複数レイテンシ列があっても、この列を優先可）

.PARAMETER ThresholdMs
  遅延のしきい値（既定: 100）

.OUTPUTS
  <Output>\TagSummary.xlsx, <Output>\TagSummary.csv

.NOTES
  - $Host などの予約語は未使用（内部は Target 名に統一）
  - パイプ要素は極力未使用（「空のパイプ要素」対策）
  - 文字化け防止のため CSV は BOM 無し UTF-8
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$CsvPath,

  [string]$Output = $null,

  [string]$TagColumn,
  [string]$TargetColumn,
  [string]$HostColumn,
  [string]$LatencyColumn,

  [int]$ThresholdMs = 100
)

function Write-Info($msg){ Write-Host "[TagSummary] $msg" }

function Get-RepoRoot(){
  # このファイルは ...\scripts\windows にある想定
  $here = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
  $scriptsDir = Split-Path -Parent -Path $here
  $repo = Split-Path -Parent -Path $scriptsDir
  return $repo
}

function Resolve-RepoPath([string]$pathIn){
  if([string]::IsNullOrWhiteSpace($pathIn)){ return $null }
  if([System.IO.Path]::IsPathRooted($pathIn)){ return $pathIn }
  $repo = Get-RepoRoot
  return (Join-Path $repo $pathIn)
}

function Ensure-Dir([string]$path){
  if([string]::IsNullOrWhiteSpace($path)){ return }
  [void][System.IO.Directory]::CreateDirectory($path)
}

function ToLowerArray($arr){
  $out = @()
  foreach($x in $arr){ $out += ($x -as [string]).ToLowerInvariant() }
  return $out
}

function Detect-TagColumn($headersLower){
  $cands = @('probe','tag','location','site','place')
  foreach($c in $cands){ if($headersLower -contains $c){ return $c } }
  return $null
}

function Detect-TargetColumn($headersLower){
  $cands = @('target','hostname','endpoint','server','dst','fqdn','ip','addr')
  foreach($c in $cands){ if($headersLower -contains $c){ return $c } }
  return $null
}

function Get-LatencyCandidates($headersOrig, $headersLower, $forced){
  # 明示指定があればその1列のみ
  if($forced){
    $idx = -1
    for($i=0;$i -lt $headersOrig.Count;$i++){
      if($headersOrig[$i] -eq $forced -or $headersLower[$i] -eq $forced.ToLowerInvariant()){ $idx = $i; break }
    }
    if($idx -ge 0){ return @(@{ Name=$headersOrig[$idx]; Lower=$headersLower[$idx]; Suffix=''; }) }
  }

  $tokens = @('rtt','latency','ping','response','delay','roundtrip','elapsed','time','avg')
  $cands = @()
  for($i=0;$i -lt $headersLower.Count;$i++){
    $h = $headersLower[$i]
    $matched = $false
    foreach($t in $tokens){ if($h.IndexOf($t) -ge 0){ $matched = $true; break } }
    if(-not $matched){ continue }

    # 列名から Target 推定
    $suffix = $h
    foreach($t in $tokens){ $suffix = $suffix.Replace($t,'') }
    $suffix = $suffix.Replace('_ms','').Replace('-ms','').Replace('ms','')
    $clean = [System.Text.RegularExpressions.Regex]::Replace($suffix, '[^a-z0-9]+', '_').Trim('_')
    $suffixGuess = ''
    if($clean -ne ''){
      $parts = $clean.Split('_')
      $suffixGuess = $parts[$parts.Length-1]
      if($suffixGuess -eq ''){ $suffixGuess = $clean }
    }
    if($suffixGuess -eq 'gw'){ $suffixGuess = 'gateway' }

    $cands += @{ Name=$headersOrig[$i]; Lower=$h; Suffix=$suffixGuess; }
  }
  return $cands
}

function Parse-DoubleOrNull([string]$s){
  if(-not $s){ return $null }
  $m = [System.Text.RegularExpressions.Regex]::Match($s, '[-+]?\d+(\.\d+)?')
  if($m.Success){ return [double]$m.Value }
  return $null
}

function Sanitize-WorksheetName([string]$name){
  if([string]::IsNullOrWhiteSpace($name)){ return "_(blank)" }
  $n = $name -replace '[\\\/\?\*\[\]:]', '_'
  if($n.Length -gt 31){ return $n.Substring(0,31) }
  return $n
}

function Percentile([double[]]$values, [double]$p){
  if(-not $values -or $values.Count -eq 0){ return $null }
  $arr = @($values); [Array]::Sort($arr)
  $n = $arr.Count
  if($n -eq 1){ return $arr[0] }
  $rank = [math]::Ceiling($p * $n)
  $idx = [math]::Min([math]::Max($rank-1,0), $n-1)
  return $arr[$idx]
}

try{
  # 絶対パスへ解決（CWD ずれ対策・OneDrive対策）
  $CsvPath = Resolve-RepoPath $CsvPath
  if(-not $Output -or $Output -eq ''){ $Output = (Join-Path (Get-RepoRoot) "Output") } else { $Output = Resolve-RepoPath $Output }

  # 入出力の存在保証
  if(-not (Test-Path -LiteralPath $CsvPath)){ throw "File not found: $CsvPath" }
  Ensure-Dir $Output

  Write-Info ("CSV: {0}" -f $CsvPath)
  Write-Info ("OUT: {0}" -f $Output)

  $rows = Import-Csv -LiteralPath $CsvPath
  if(-not $rows){ throw "CSV rows are empty." }

  # ヘッダ
  $headersOrig = @($rows[0].PSObject.Properties.Name)
  $headersLower = ToLowerArray $headersOrig

  # 列検出
  $tagColLower = if($TagColumn){ $TagColumn.ToLowerInvariant() } else { Detect-TagColumn $headersLower }
  $tgtColLower = $null
  if($TargetColumn){ $tgtColLower = $TargetColumn.ToLowerInvariant() }
  elseif($HostColumn){ $tgtColLower = $HostColumn.ToLowerInvariant() }
  else { $tgtColLower = Detect-TargetColumn $headersLower }

  # 実ヘッダ名へ引き直し
  $tagCol = $null; $tgtCol = $null
  for($i=0;$i -lt $headersOrig.Count;$i++){
    if($tagColLower -and $headersLower[$i] -eq $tagColLower){ $tagCol = $headersOrig[$i] }
    if($tgtColLower -and $headersLower[$i] -eq $tgtColLower){ $tgtCol = $headersOrig[$i] }
  }

  # レイテンシ候補
  $latCands = Get-LatencyCandidates $headersOrig $headersLower $LatencyColumn
  if(-not $latCands -or $latCands.Count -eq 0){
    throw "レイテンシ列が見つかりません。列名に rtt/latency/ping/response/delay/roundtrip/elapsed/time/avg を含めてください。"
  }

  Write-Info ("Latency columns detected: {0}" -f $latCands.Count)

  # ロング化
  $proj = @()
  if($latCands.Count -eq 1){
    $latName = $latCands[0].Name
    foreach($r in $rows){
      $tag = if($tagCol){ ($r.$tagCol -as [string]) } else { $null }
      $tgt = if($tgtCol){ ($r.$tgtCol -as [string]) } else { "_(all)" }
      $lat = Parse-DoubleOrNull ($r.$latName -as [string])
      if($lat -ne $null){
        $proj += [pscustomobject]@{
          Tag = if([string]::IsNullOrWhiteSpace($tag)){"_(blank)"}else{$tag}
          Target = if([string]::IsNullOrWhiteSpace($tgt)){"_(all)"}else{$tgt}
          LatencyMs = $lat
        }
      }
    }
  } else {
    foreach($r in $rows){
      $tag = if($tagCol){ ($r.$tagCol -as [string]) } else { $null }
      foreach($c in $latCands){
        $lat = Parse-DoubleOrNull ($r.($c.Name) -as [string])
        if($lat -eq $null){ continue }
        $tgtFromHeader = $c.Suffix
        if(-not $tgtFromHeader -or $tgtFromHeader -eq ''){ $tgtFromHeader = $c.Name }
        $proj += [pscustomobject]@{
          Tag = if([string]::IsNullOrWhiteSpace($tag)){"_(blank)"}else{$tag}
          Target = $tgtFromHeader
          LatencyMs = $lat
        }
      }
    }
  }

  if(-not $proj -or $proj.Count -eq 0){ throw "有効なレイテンシ値が1件もありません。" }

  # 手動グルーピング（Tag, Target）
  $dict = @{}
  foreach($p in $proj){
    $key = "$($p.Tag)`t$($p.Target)"
    if(-not $dict.ContainsKey($key)){
      $dict[$key] = [pscustomobject]@{ Tag=$p.Tag; Target=$p.Target; Values=@() }
    }
    $vals = @($dict[$key].Values); $vals += $p.LatencyMs; $dict[$key].Values = $vals
  }

  # 集計
  $outRows = @()
  foreach($kv in $dict.GetEnumerator()){
    $rec = $kv.Value
    $vals = @($rec.Values)
    $n = $vals.Count

    # 平均
    $sum = 0.0; foreach($v in $vals){ $sum += [double]$v }
    $avg = if($n -gt 0){ [math]::Round($sum / $n, 1) } else { $null }

    # P95
    $p95 = [math]::Round((Percentile $vals 0.95), 1)

    # 閾値超え
    $over = 0; foreach($v in $vals){ if($v -ge $ThresholdMs){ $over++ } }
    $overPct = if($n -gt 0){ [math]::Round(100.0 * $over / $n, 1) } else { $null }

    $outRows += [pscustomobject]@{
      Tag=$rec.Tag; Target=$rec.Target; Samples=$n; AvgMs=$avg; P95Ms=$p95;
      OverThresholdCount=$over; OverThresholdPct=$overPct; ThresholdMs=$ThresholdMs
    }
  }

  # 並び替え（Tag, Target）
  $sorted = @($outRows)
  for($i=0; $i -lt $sorted.Count; $i++){
    for($j=$i+1; $j -lt $sorted.Count; $j++){
      $a = $sorted[$i]; $b = $sorted[$j]
      $cmpTag = [string]::Compare($a.Tag, $b.Tag)
      $swap = $false
      if($cmpTag -gt 0){ $swap = $true }
      elseif($cmpTag -eq 0){ if([string]::Compare($a.Target, $b.Target) -gt 0){ $swap = $true } }
      if($swap){ $tmp=$sorted[$i]; $sorted[$i]=$sorted[$j]; $sorted[$j]=$tmp }
    }
  }

  # CSV 出力（BOM 無し UTF-8）
  $csvOut = Join-Path $Output "TagSummary.csv"
  Ensure-Dir (Split-Path -Parent -Path $csvOut)
  Write-Info ("Write CSV: {0}" -f $csvOut)
  $lines = @()
  $lines += '"Tag","Target","Samples","AvgMs","P95Ms","OverThresholdCount","OverThresholdPct","ThresholdMs"'
  foreach($r in $sorted){
    $line = '"' + ($r.Tag -replace '"','""') + '","' +
                  ($r.Target -replace '"','""') + '",' +
                  $r.Samples + ',' + $r.AvgMs + ',' + $r.P95Ms + ',' +
                  $r.OverThresholdCount + ',' + $r.OverThresholdPct + ',' + $r.ThresholdMs
    $lines += $line
  }
  [System.IO.File]::WriteAllLines($csvOut, $lines, (New-Object System.Text.UTF8Encoding($false)))

  # Excel 出力（Office 環境のみ）
  $xlsxOut = Join-Path $Output "TagSummary.xlsx"
  Ensure-Dir (Split-Path -Parent -Path $xlsxOut)
  $excel = $null
  try{ $excel = New-Object -ComObject Excel.Application }catch{}
  if($excel){
    try{
      $excel.Visible = $false
      $wb = $excel.Workbooks.Add()
      $headers = @('Tag','Target','Samples','AvgMs','P95Ms','OverThresholdCount','OverThresholdPct','ThresholdMs')

      # Summary
      $ws = $wb.Worksheets.Item(1); $ws.Name = "Summary"
      for($i=0;$i -lt $headers.Count;$i++){ $ws.Cells.Item(1, $i+1) = $headers[$i] }
      $row = 2
      foreach($r in $sorted){
        $ws.Cells.Item($row,1)=$r.Tag; $ws.Cells.Item($row,2)=$r.Target; $ws.Cells.Item($row,3)=$r.Samples
        $ws.Cells.Item($row,4)=$r.AvgMs; $ws.Cells.Item($row,5)=$r.P95Ms
        $ws.Cells.Item($row,6)=$r.OverThresholdCount; $ws.Cells.Item($row,7)=$r.OverThresholdPct
        $ws.Cells.Item($row,8)=$r.ThresholdMs; $row++
      }

      # Tag 毎
      $byTag = @{}
      foreach($r in $sorted){ if(-not $byTag.ContainsKey($r.Tag)){ $byTag[$r.Tag]=@() }; $byTag[$r.Tag]+=$r }
      foreach($kv in $byTag.GetEnumerator()){
        $name = Sanitize-WorksheetName $kv.Key
        $sheet = $wb.Worksheets.Add(); $null = $sheet.Move($ws); $sheet.Name = $name
        for($i=0;$i -lt $headers.Count;$i++){ $sheet.Cells.Item(1, $i+1) = $headers[$i] }
        # Target 昇順
        $grp = @($kv.Value)
        for($i=0;$i -lt $grp.Count;$i++){ for($j=$i+1;$j -lt $grp.Count;$j++){
          if([string]::Compare($grp[$i].Target,$grp[$j].Target) -gt 0){ $t=$grp[$i];$grp[$i]=$grp[$j];$grp[$j]=$t }
        } }
        $ridx=2
        foreach($r in $grp){
          $sheet.Cells.Item($ridx,1)=$r.Tag; $sheet.Cells.Item($ridx,2)=$r.Target; $sheet.Cells.Item($ridx,3)=$r.Samples
          $sheet.Cells.Item($ridx,4)=$r.AvgMs; $sheet.Cells.Item($ridx,5)=$r.P95Ms
          $sheet.Cells.Item($ridx,6)=$r.OverThresholdCount; $sheet.Cells.Item($ridx,7)=$r.OverThresholdPct
          $sheet.Cells.Item($ridx,8)=$r.ThresholdMs; $ridx++
        }
      }

      # 体裁
      $cnt = $wb.Worksheets.Count
      for($i=1;$i -le $cnt;$i++){ $s=$wb.Worksheets.Item($i); $s.Rows.Item(1).Font.Bold=$true; $null=$s.Columns.AutoFit() }

      $wb.SaveAs($xlsxOut); $wb.Close($true)
    } finally {
      $null = $excel.Quit()
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
  }

  Write-Info "Done."
}
catch{
  Write-Error $_.Exception.Message
  exit 1
}
