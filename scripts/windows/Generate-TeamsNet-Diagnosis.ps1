<#
.SYNOPSIS
  エリア(Tag) × AP(SSID/BSSID/AP名) × Target 別のRTT診断、および時間帯(Hour)別の傾向を集計。
  PowerShell 5.1対応・OneDrive/日本語パス対応。Run-Merge-TeamsNet-CSVの出力に整合。

.DESCRIPTION
  - Latency列（例: rtt_ms_gateway / rtt_ms_hop2 / rtt_ms_hop3 など横持ち）を自動検出。
    複数あれば「列名の接尾辞→Target」に読み替えてロング化。
    1列だけなら TargetColumn（または HostColumn）を使ってTargetを取得。
  - floors.csv で BSSID / AP名 による floor/area を付与（任意）。
  - 出力:
      <Output>\TeamsNet_Diagnosis.xlsx
        - SummaryWorst …… Slow%の高い組み合わせ上位（Tag×AP×Target）
        - ByArea_AP_Target … エリア(Tag)×AP×Target の全集計
        - TimeOfDay … Tag×Hour×Target の時間帯傾向
      <Output>\Diag_ByArea_AP_Target.csv
      <Output>\Diag_TimeOfDay.csv

.PARAMETER CsvPath
  入力CSV。相対ならリポジトリ直下を基準に解決。

.PARAMETER Output
  出力フォルダ（既定: <repo>\Output）。相対ならリポジトリ直下基準。

.PARAMETER FloorMap
  floors.csv のパス（bssid/ap_name/floor/area など）。未指定かつ <repo>\floors.csv が存在すれば自動使用。

.PARAMETER BucketMinutes
  時間帯集計のバケット幅（分）。既定 60（=時間帯別）。15などにすると15分単位。

.PARAMETER ThresholdMs
  RTTの遅延しきい値（既定 100）。Slow%＝しきい値以上の割合。

※ 列自動検出
  Tag: probe/tag/location/site/place
  Target: target/hostname/endpoint/server/dst/fqdn/ip/addr
  Time: timestamp/time/datetime/local_time/utc/created_at/measured_at
  AP: ssid/bssid/ap_name

#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$CsvPath,

  [string]$Output = $null,
  [string]$FloorMap = $null,

  [int]$BucketMinutes = 60,
  [int]$ThresholdMs = 100,

  [string]$TagColumn,
  [string]$TargetColumn,
  [string]$HostColumn,
  [string]$LatencyColumn,
  [string]$TimeColumn,
  [string]$BssidColumn,
  [string]$ApNameColumn,
  [string]$SsidColumn
)

# ---- 共通ユーティリティ ----
function Write-Info($m){ Write-Host "[Diagnosis] $m" }

function Get-RepoRoot(){
  $here = Split-Path -Parent -Path $MyInvocation.MyCommand.Path   # ...\scripts\windows
  $scripts = Split-Path -Parent -Path $here                        # ...\scripts
  $repo = Split-Path -Parent -Path $scripts                        # <repo>
  return $repo
}

function Resolve-RepoPath([string]$p){
  if([string]::IsNullOrWhiteSpace($p)){ return $null }
  if([System.IO.Path]::IsPathRooted($p)){ return $p }
  return (Join-Path (Get-RepoRoot) $p)
}

function Ensure-Dir([string]$path){
  if([string]::IsNullOrWhiteSpace($path)){ return }
  [void][System.IO.Directory]::CreateDirectory($path)
}

function ToLowerArray($arr){
  $o=@(); foreach($x in $arr){ $o += ($x -as [string]).ToLowerInvariant() }; return $o
}

function Detect-First($headersLower, $cands){
  foreach($c in $cands){ if($headersLower -contains $c){ return $c } }
  return $null
}

function NormalizeBssid([string]$s){
  if(-not $s){ return $null }
  $hex = [System.Text.RegularExpressions.Regex]::Replace($s.ToLowerInvariant(), '[^0-9a-f]', '')
  if($hex.Length -ne 12){ return $null }
  return $hex
}

function Parse-DoubleOrNull([string]$s){
  if(-not $s){ return $null }
  $m = [System.Text.RegularExpressions.Regex]::Match($s, '[-+]?\d+(\.\d+)?')
  if($m.Success){ return [double]$m.Value }
  return $null
}

function Parse-DateOrNull([string]$s){
  if(-not $s){ return $null }
  try{ return [datetime]::Parse($s) }catch{ return $null }
}

function RoundDownMinutes([datetime]$dt, [int]$mins){
  if(-not $dt){ return $null }
  if($mins -lt 1){ $mins = 60 }
  $m = $dt.Minute - ($dt.Minute % $mins)
  return (Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour $dt.Hour -Minute $m -Second 0)
}

function Percentile([double[]]$values, [double]$p){
  if(-not $values -or $values.Count -eq 0){ return $null }
  $arr=@($values); [Array]::Sort($arr)
  $n=$arr.Count; if($n -eq 1){ return $arr[0] }
  $rank=[math]::Ceiling($p*$n); $idx=[math]::Min([math]::Max($rank-1,0),$n-1)
  return $arr[$idx]
}

function Sanitize-WorksheetName([string]$name){
  if([string]::IsNullOrWhiteSpace($name)){ return "_(blank)" }
  $n=$name -replace '[\\\/\?\*\[\]:]', '_'; if($n.Length -gt 31){ return $n.Substring(0,31) }; return $n
}

# ---- メイン処理 ----
try{
  # パス解決
  $CsvPath = Resolve-RepoPath $CsvPath
  if(-not $Output -or $Output -eq ''){ $Output = (Join-Path (Get-RepoRoot) "Output") } else { $Output = Resolve-RepoPath $Output }
  if(-not $FloorMap -or $FloorMap -eq ''){
    $autoFloors = Join-Path (Get-RepoRoot) "floors.csv"
    if(Test-Path -LiteralPath $autoFloors){ $FloorMap = $autoFloors }
  } else { $FloorMap = Resolve-RepoPath $FloorMap }

  if(-not (Test-Path -LiteralPath $CsvPath)){ throw "File not found: $CsvPath" }
  Ensure-Dir $Output

  Write-Info ("CSV: {0}" -f $CsvPath)
  if($FloorMap){ Write-Info ("FloorMap: {0}" -f $FloorMap) }
  Write-Info ("OUT: {0}" -f $Output)

  $rows = Import-Csv -LiteralPath $CsvPath
  if(-not $rows){ throw "CSV rows are empty." }

  # ヘッダ
  $headersOrig = @($rows[0].PSObject.Properties.Name)
  $headersLower = ToLowerArray $headersOrig

  # 列検出
  $tagLower = if($TagColumn){ $TagColumn.ToLowerInvariant() } else { Detect-First $headersLower @('probe','tag','location','site','place') }
  $tgtLower = $null
  if($TargetColumn){ $tgtLower = $TargetColumn.ToLowerInvariant() }
  elseif($HostColumn){ $tgtLower = $HostColumn.ToLowerInvariant() }
  else { $tgtLower = Detect-First $headersLower @('target','hostname','endpoint','server','dst','fqdn','ip','addr') }

  $timeLower = if($TimeColumn){ $TimeColumn.ToLowerInvariant() } else { Detect-First $headersLower @('timestamp','time','datetime','local_time','utc','created_at','measured_at') }
  $bssidLower = if($BssidColumn){ $BssidColumn.ToLowerInvariant() } else { Detect-First $headersLower @('bssid','ap_bssid') }
  $apnameLower = if($ApNameColumn){ $ApNameColumn.ToLowerInvariant() } else { Detect-First $headersLower @('ap_name','apname','ap') }
  $ssidLower = if($SsidColumn){ $SsidColumn.ToLowerInvariant() } else { Detect-First $headersLower @('ssid','wifi_ssid') }

  # 実列名へ引き直し
  $tagCol=$null;$tgtCol=$null;$timeCol=$null;$bssidCol=$null;$apnameCol=$null;$ssidCol=$null
  for($i=0;$i -lt $headersOrig.Count;$i++){
    if($tagLower -and $headersLower[$i] -eq $tagLower){ $tagCol = $headersOrig[$i] }
    if($tgtLower -and $headersLower[$i] -eq $tgtLower){ $tgtCol = $headersOrig[$i] }
    if($timeLower -and $headersLower[$i] -eq $timeLower){ $timeCol = $headersOrig[$i] }
    if($bssidLower -and $headersLower[$i] -eq $bssidLower){ $bssidCol = $headersOrig[$i] }
    if($apnameLower -and $headersLower[$i] -eq $apnameLower){ $apnameCol = $headersOrig[$i] }
    if($ssidLower -and $headersLower[$i] -eq $ssidLower){ $ssidCol = $headersOrig[$i] }
  }

  # Latency候補の抽出（横持ち対応）
  function Get-LatencyCandidates($headersOrig, $headersLower, $forced){
    if($forced){
      for($i=0;$i -lt $headersOrig.Count;$i++){
        if($headersOrig[$i] -eq $forced -or $headersLower[$i] -eq $forced.ToLowerInvariant()){
          return @(@{ Name=$headersOrig[$i]; Lower=$headersLower[$i]; Suffix='' })
        }
      }
    }
    $tokens=@('rtt','latency','ping','response','delay','roundtrip','elapsed','time','avg')
    $c=@()
    for($i=0;$i -lt $headersLower.Count;$i++){
      $h=$headersLower[$i]; $ok=$false
      foreach($t in $tokens){ if($h.IndexOf($t) -ge 0){ $ok=$true; break } }
      if(-not $ok){ continue }
      $suffix=$h; foreach($t in $tokens){ $suffix=$suffix.Replace($t,'') }
      $suffix=$suffix.Replace('_ms','').Replace('-ms','').Replace('ms','')
      $clean=[System.Text.RegularExpressions.Regex]::Replace($suffix,'[^a-z0-9]+','_').Trim('_')
      $guess=''; if($clean -ne ''){ $parts=$clean.Split('_'); $guess=$parts[$parts.Length-1]; if($guess -eq ''){ $guess=$clean } }
      if($guess -eq 'gw'){ $guess='gateway' }
      $c += @{ Name=$headersOrig[$i]; Lower=$h; Suffix=$guess }
    }
    return $c
  }

  $latCands = Get-LatencyCandidates $headersOrig $headersLower $LatencyColumn
  if(-not $latCands -or $latCands.Count -eq 0){ throw "レイテンシ列が見つかりません。" }

  # floors.csv の読み込み（任意）
  $floorByBssid=@{}; $floorByApName=@{}; $areaByBssid=@{}; $areaByApName=@{}
  if($FloorMap -and (Test-Path -LiteralPath $FloorMap)){
    $fm = Import-Csv -LiteralPath $FloorMap
    foreach($f in $fm){
      $bssidRaw = $null; if($f.PSObject.Properties['bssid']){ $bssidRaw = ($f.'bssid' -as [string]) }
      $apnRaw = $null; if($f.PSObject.Properties['ap_name']){ $apnRaw = ($f.'ap_name' -as [string]) }
      $floorVal = $null; if($f.PSObject.Properties['floor']){ $floorVal = ($f.'floor' -as [string]) }
      $areaVal = $null; if($f.PSObject.Properties['area']){ $areaVal = ($f.'area' -as [string]) }

      $nb = NormalizeBssid $bssidRaw
      if($nb){
        if($floorVal){ $floorByBssid[$nb] = $floorVal }
        if($areaVal){ $areaByBssid[$nb] = $areaVal }
      }
      if($apnRaw){
        $key = $apnRaw.ToLowerInvariant()
        if($floorVal){ $floorByApName[$key] = $floorVal }
        if($areaVal){ $areaByApName[$key] = $areaVal }
      }
    }
  }

  # ---- ロング化（Tag/Target/LatencyMs + AP/Time 情報付き） ----
  $proj=@()
  if($latCands.Count -eq 1){
    $latName=$latCands[0].Name
    foreach($r in $rows){
      $tag = if($tagCol){ ($r.$tagCol -as [string]) } else { $null }
      $tgt = if($tgtCol){ ($r.$tgtCol -as [string]) } else { "_(all)" }
      $lat = Parse-DoubleOrNull ($r.$latName -as [string]); if($lat -eq $null){ continue }

      $ssid = if($ssidCol){ ($r.$ssidCol -as [string]) } else { $null }
      $apn  = if($apnameCol){ ($r.$apnameCol -as [string]) } else { $null }
      $bss  = if($bssidCol){ ($r.$bssidCol -as [string]) } else { $null }
      $nb   = NormalizeBssid $bss

      $dt = $null; if($timeCol){ $dt = Parse-DateOrNull ($r.$timeCol -as [string]) }
      $bucket = if($dt){ RoundDownMinutes $dt $BucketMinutes } else { $null }
      $hour = if($dt){ $dt.Hour } else { $null }

      $proj += [pscustomobject]@{
        Tag = if([string]::IsNullOrWhiteSpace($tag)){"_(blank)"}else{$tag}
        Target = if([string]::IsNullOrWhiteSpace($tgt)){"_(all)"}else{$tgt}
        LatencyMs = $lat
        SSID = $ssid
        APName = $apn
        BSSID = $bss
        NBSSID = $nb
        Floor = $null
        Area = $null
        Time = $dt
        Hour = $hour
        Bucket = $bucket
      }
    }
  } else {
    foreach($r in $rows){
      $tag = if($tagCol){ ($r.$tagCol -as [string]) } else { $null }
      $ssid = if($ssidCol){ ($r.$ssidCol -as [string]) } else { $null }
      $apn  = if($apnameCol){ ($r.$apnameCol -as [string]) } else { $null }
      $bss  = if($bssidCol){ ($r.$bssidCol -as [string]) } else { $null }
      $nb   = NormalizeBssid $bss

      $dt = $null; if($timeCol){ $dt = Parse-DateOrNull ($r.$timeCol -as [string]) }
      $bucket = if($dt){ RoundDownMinutes $dt $BucketMinutes } else { $null }
      $hour = if($dt){ $dt.Hour } else { $null }

      foreach($c in $latCands){
        $lat = Parse-DoubleOrNull ($r.($c.Name) -as [string]); if($lat -eq $null){ continue }
        $tgt = $c.Suffix; if(-not $tgt -or $tgt -eq ''){ $tgt = $c.Name }
        $proj += [pscustomobject]@{
          Tag = if([string]::IsNullOrWhiteSpace($tag)){"_(blank)"}else{$tag}
          Target = $tgt
          LatencyMs = $lat
          SSID = $ssid
          APName = $apn
          BSSID = $bss
          NBSSID = $nb
          Floor = $null
          Area = $null
          Time = $dt
          Hour = $hour
          Bucket = $bucket
        }
      }
    }
  }

  if(-not $proj -or $proj.Count -eq 0){ throw "有効なレイテンシ値が1件もありません。" }

  # floors.csv マージ（BSSID優先→AP名）
  if($FloorMap){
    for($i=0;$i -lt $proj.Count;$i++){
      $nb=$proj[$i].NBSSID; $apnKey=$null
      if($proj[$i].APName){ $apnKey = $proj[$i].APName.ToLowerInvariant() }
      $floor=$null;$area=$null
      if($nb -and $floorByBssid.ContainsKey($nb)){ $floor=$floorByBssid[$nb] }
      if(-not $floor -and $apnKey -and $floorByApName.ContainsKey($apnKey)){ $floor=$floorByApName[$apnKey] }
      if($nb -and $areaByBssid.ContainsKey($nb)){ $area=$areaByBssid[$nb] }
      if(-not $area -and $apnKey -and $areaByApName.ContainsKey($apnKey)){ $area=$areaByApName[$apnKey] }
      $proj[$i].Floor = if($floor){ $floor } else { "Unknown" }
      $proj[$i].Area  = if($area){ $area } else { $null }
    }
  }

  # ---- 集計: ByArea_AP_Target (Tag×AP×Target) ----
  $dict=@{}  # key: Tag\tAPName\tBSSID\tSSID\tFloor\tTarget
  foreach($p in $proj){
    $apDsp = $p.APName; if(-not $apDsp -or $apDsp -eq ''){ $apDsp = "(AP unknown)" }
    $bssidDsp = if($p.BSSID){ $p.BSSID } else { "" }
    $floorDsp = if($p.Floor){ $p.Floor } else { "Unknown" }
    $ssidDsp = if($p.SSID){ $p.SSID } else { "" }

    $key = "$($p.Tag)`t$apDsp`t$bssidDsp`t$ssidDsp`t$floorDsp`t$($p.Target)"
    if(-not $dict.ContainsKey($key)){
      $dict[$key] = [pscustomobject]@{
        Tag=$p.Tag; APName=$apDsp; BSSID=$bssidDsp; SSID=$ssidDsp; Floor=$floorDsp; Target=$p.Target;
        Values=@()
      }
    }
    $vals=@($dict[$key].Values); $vals += $p.LatencyMs; $dict[$key].Values=$vals
  }

  $byArea=@()
  foreach($kv in $dict.GetEnumerator()){
    $rec=$kv.Value; $vals=@($rec.Values); $n=$vals.Count
    $sum=0.0; foreach($v in $vals){ $sum += [double]$v }
    $avg = if($n -gt 0){ [math]::Round($sum/$n,1) } else { $null }
    $p95 = [math]::Round((Percentile $vals 0.95),1)
    $over=0; foreach($v in $vals){ if($v -ge $ThresholdMs){ $over++ } }
    $overPct = if($n -gt 0){ [math]::Round(100.0*$over/$n,1) } else { $null }

    $byArea += [pscustomobject]@{
      Tag=$rec.Tag; APName=$rec.APName; BSSID=$rec.BSSID; SSID=$rec.SSID; Floor=$rec.Floor; Target=$rec.Target;
      Samples=$n; AvgMs=$avg; P95Ms=$p95; OverThresholdCount=$over; OverThresholdPct=$overPct; ThresholdMs=$ThresholdMs
    }
  }

  # ---- 集計: TimeOfDay (Tag×Hour×Target) ----
  $td=@{}  # key: Tag\tHour\tTarget
  foreach($p in $proj){
    $hour = if($p.Hour -ne $null){ $p.Hour } else { -1 }
    $key="$($p.Tag)`t$hour`t$($p.Target)"
    if(-not $td.ContainsKey($key)){
      $td[$key] = [pscustomobject]@{ Tag=$p.Tag; Hour=$hour; Target=$p.Target; Values=@() }
    }
    $vals=@($td[$key].Values); $vals+=$p.LatencyMs; $td[$key].Values=$vals
  }

  $byTime=@()
  foreach($kv in $td.GetEnumerator()){
    $rec=$kv.Value; $vals=@($rec.Values); $n=$vals.Count
    $sum=0.0; foreach($v in $vals){ $sum += [double]$v }
    $avg = if($n -gt 0){ [math]::Round($sum/$n,1) } else { $null }
    $p95 = [math]::Round((Percentile $vals 0.95),1)
    $over=0; foreach($v in $vals){ if($v -ge $ThresholdMs){ $over++ } }
    $overPct = if($n -gt 0){ [math]::Round(100.0*$over/$n,1) } else { $null }
    $byTime += [pscustomobject]@{
      Tag=$rec.Tag; Hour=$rec.Hour; Target=$rec.Target; Samples=$n; AvgMs=$avg; P95Ms=$p95;
      OverThresholdCount=$over; OverThresholdPct=$overPct; ThresholdMs=$ThresholdMs
    }
  }

  # ---- 並び替え（安定化） ----
  function Sort-ByStringPairs($arr, $keys){
    $sorted=@($arr)
    for($i=0;$i -lt $sorted.Count;$i++){
      for($j=$i+1;$j -lt $sorted.Count;$j++){
        $a=$sorted[$i]; $b=$sorted[$j]; $swap=$false
        foreach($k in $keys){
          $ka=$a.$k; $kb=$b.$k
          $cmp=[string]::Compare(("$ka"),("$kb"))
          if($cmp -gt 0){ $swap=$true; break } elseif($cmp -lt 0){ break }
        }
        if($swap){ $t=$sorted[$i]; $sorted[$i]=$sorted[$j]; $sorted[$j]=$t }
      }
    }
    return ,$sorted
  }

  $byAreaSorted = Sort-ByStringPairs $byArea @('Tag','APName','Target')
  $byTimeSorted = Sort-ByStringPairs $byTime @('Tag','Hour','Target')

  # ---- CSV書き出し（BOMなしUTF-8） ----
  $csv1 = Join-Path $Output "Diag_ByArea_AP_Target.csv"
  $csv2 = Join-Path $Output "Diag_TimeOfDay.csv"
  Ensure-Dir (Split-Path -Parent -Path $csv1)
  Ensure-Dir (Split-Path -Parent -Path $csv2)

  $lines=@()
  $lines += '"Tag","APName","BSSID","SSID","Floor","Target","Samples","AvgMs","P95Ms","OverThresholdCount","OverThresholdPct","ThresholdMs"'
  foreach($r in $byAreaSorted){
    $lines += '"' + ($r.Tag -replace '"','""') + '","' + ($r.APName -replace '"','""') + '","' +
              ($r.BSSID -replace '"','""') + '","' + ($r.SSID -replace '"','""') + '","' +
              ($r.Floor -replace '"','""') + '","' + ($r.Target -replace '"','""') + '",' +
              $r.Samples + ',' + $r.AvgMs + ',' + $r.P95Ms + ',' +
              $r.OverThresholdCount + ',' + $r.OverThresholdPct + ',' + $r.ThresholdMs
  }
  [System.IO.File]::WriteAllLines($csv1, $lines, (New-Object System.Text.UTF8Encoding($false)))

  $lines=@()
  $lines += '"Tag","Hour","Target","Samples","AvgMs","P95Ms","OverThresholdCount","OverThresholdPct","ThresholdMs"'
  foreach($r in $byTimeSorted){
    $lines += '"' + ($r.Tag -replace '"','""') + '",' + $r.Hour + ',"' + ($r.Target -replace '"','""') + '",' +
              $r.Samples + ',' + $r.AvgMs + ',' + $r.P95Ms + ',' +
              $r.OverThresholdCount + ',' + $r.OverThresholdPct + ',' + $r.ThresholdMs
  }
  [System.IO.File]::WriteAllLines($csv2, $lines, (New-Object System.Text.UTF8Encoding($false)))

  # ---- Excel 出力（任意） ----
  $xlsx = Join-Path $Output "TeamsNet_Diagnosis.xlsx"
  $excel=$null; try{ $excel=New-Object -ComObject Excel.Application }catch{}
  if($excel){
    try{
      $excel.Visible=$false
      $wb=$excel.Workbooks.Add()
      $headers1=@('Tag','APName','BSSID','SSID','Floor','Target','Samples','AvgMs','P95Ms','OverThresholdCount','OverThresholdPct','ThresholdMs')
      $headers2=@('Tag','Hour','Target','Samples','AvgMs','P95Ms','OverThresholdCount','OverThresholdPct','ThresholdMs')

      # SummaryWorst（Slow%降順 上位50）
      $ws=$wb.Worksheets.Item(1); $ws.Name="SummaryWorst"
      for($i=0;$i -lt $headers1.Count;$i++){ $ws.Cells.Item(1,$i+1)=$headers1[$i] }
      # 手動ソート: by OverThresholdPct desc
      $top=@($byAreaSorted)
      # バブル降順
      for($i=0;$i -lt $top.Count;$i++){
        for($j=$i+1;$j -lt $top.Count;$j++){
          if($top[$i].OverThresholdPct -lt $top[$j].OverThresholdPct){
            $t=$top[$i]; $top[$i]=$top[$j]; $top[$j]=$t
          }
        }
      }
      $maxRows = if($top.Count -gt 50){ 50 } else { $top.Count }
      $row=2
      for($k=0;$k -lt $maxRows;$k++){
        $r=$top[$k]
        $ws.Cells.Item($row,1)=$r.Tag; $ws.Cells.Item($row,2)=$r.APName; $ws.Cells.Item($row,3)=$r.BSSID
        $ws.Cells.Item($row,4)=$r.SSID; $ws.Cells.Item($row,5)=$r.Floor; $ws.Cells.Item($row,6)=$r.Target
        $ws.Cells.Item($row,7)=$r.Samples; $ws.Cells.Item($row,8)=$r.AvgMs; $ws.Cells.Item($row,9)=$r.P95Ms
        $ws.Cells.Item($row,10)=$r.OverThresholdCount; $ws.Cells.Item($row,11)=$r.OverThresholdPct; $ws.Cells.Item($row,12)=$r.ThresholdMs
        $row++
      }

      # ByArea_AP_Target
      $sheet=$wb.Worksheets.Add(); $null=$sheet.Move($ws); $sheet.Name="ByArea_AP_Target"
      for($i=0;$i -lt $headers1.Count;$i++){ $sheet.Cells.Item(1,$i+1)=$headers1[$i] }
      $row=2
      foreach($r in $byAreaSorted){
        $sheet.Cells.Item($row,1)=$r.Tag; $sheet.Cells.Item($row,2)=$r.APName; $sheet.Cells.Item($row,3)=$r.BSSID
        $sheet.Cells.Item($row,4)=$r.SSID; $sheet.Cells.Item($row,5)=$r.Floor; $sheet.Cells.Item($row,6)=$r.Target
        $sheet.Cells.Item($row,7)=$r.Samples; $sheet.Cells.Item($row,8)=$r.AvgMs; $sheet.Cells.Item($row,9)=$r.P95Ms
        $sheet.Cells.Item($row,10)=$r.OverThresholdCount; $sheet.Cells.Item($row,11)=$r.OverThresholdPct; $sheet.Cells.Item($row,12)=$r.ThresholdMs
        $row++
      }

      # TimeOfDay
      $sheet=$wb.Worksheets.Add(); $null=$sheet.Move($ws); $sheet.Name="TimeOfDay"
      for($i=0;$i -lt $headers2.Count;$i++){ $sheet.Cells.Item(1,$i+1)=$headers2[$i] }
      $row=2
      foreach($r in $byTimeSorted){
        $sheet.Cells.Item($row,1)=$r.Tag; $sheet.Cells.Item($row,2)=$r.Hour; $sheet.Cells.Item($row,3)=$r.Target
        $sheet.Cells.Item($row,4)=$r.Samples; $sheet.Cells.Item($row,5)=$r.AvgMs; $sheet.Cells.Item($row,6)=$r.P95Ms
        $sheet.Cells.Item($row,7)=$r.OverThresholdCount; $sheet.Cells.Item($row,8)=$r.OverThresholdPct; $sheet.Cells.Item($row,9)=$r.ThresholdMs
        $row++
      }

      # 体裁
      for($i=1;$i -le $wb.Worksheets.Count;$i++){
        $s=$wb.Worksheets.Item($i)
        $s.Rows.Item(1).Font.Bold=$true
        $null=$s.Columns.AutoFit()
      }

      $wb.SaveAs($xlsx); $wb.Close($true)
    } finally {
      $null=$excel.Quit()
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
  }

  Write-Info "Done."
}
catch{
  Write-Error $_.Exception.Message
  exit 1
}
