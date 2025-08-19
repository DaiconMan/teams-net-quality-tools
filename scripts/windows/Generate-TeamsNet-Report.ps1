<#
Generate-TeamsNet-Report.ps1

- 既存の teams_net_quality.csv を集計し、同一ブックに以下を出力:
  1) Host* シート: 各ターゲット(=targets.csvの行)ごとに1シート。フロア(8F/10F等)別に系列を分けて色分けした時系列グラフを作成
     - SAAS 等でICMP不可でも、役割別に ICMP/TCP/HTTP の優先順位で「有効RTT(eff_rtt_ms)」を採用
     - 閾値線(既定 100ms, 赤破線)、Y軸 0..300ms, X軸は1時間刻み表示
  2) LayerSeries/DeltaSeries: L2/L3/RTR_LAN/RTR_WAN/ZSCALER/SAAS の時間バケット平均と、区間差(Δ)のグラフ

- PS 5.1 互換。Excelは成功/失敗に関わらず必ず Quit/Release
- フロア推定: ap_name/bssid/ssid に "8F/8階/10F/10階" などが含まれればその値。なければ Unknown
- 追加の明示マップがある場合は -FloorMap で bssid,ap_name,floor を与えると優先使用

使い方:
  powershell -NoProfile -ExecutionPolicy Bypass `
    -File .\Generate-TeamsNet-Report.ps1 `
    -CsvPath "$Env:LOCALAPPDATA\TeamsNet\teams_net_quality.csv" `
    -TargetsCsv ".\targets.csv" `
    -Output ".\TeamsNet-Report.xlsx" `
    -BucketMinutes 5 `
    -ThresholdMs 100 `
    [-FloorMap .\floors.csv] [-Visible]
#>

param(
  [Parameter(Mandatory=$true)][string]$CsvPath,
  [Parameter(Mandatory=$true)][string]$TargetsCsv,
  [Parameter(Mandatory=$true)][string]$Output,
  [int]$BucketMinutes = 5,
  [int]$ThresholdMs = 100,
  [string]$FloorMap,
  [switch]$Visible
)

# ---------- 共通 ----------
$ErrorActionPreference='Stop'
$PSDefaultParameterValues['*:ErrorAction']='Stop'

function Release-Com([object]$obj){
  if($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)){
    try{ [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }catch{}
  }
}
function Sanitize-SheetName([string]$name){
  if(-not $name){ return 'Sheet' }
  $n = $name -replace '[:\\/\?\*$begin:math:display$$end:math:display$]','_'
  if($n.Length -gt 31){ $n=$n.Substring(0,31) }
  if($n -match '^\s*$'){ $n='Sheet' }
  return $n
}
function Normalize-Host([string]$s){
  if(-not $s){ return '' }
  $t = $s.Trim().Trim('"',"'").ToLowerInvariant()
  if($t -match '$begin:math:text$([0-9]{1,3}(?:\\.[0-9]{1,3}){3})$end:math:text$'){ return $Matches[1] }  # name (ip)
  if($t -match '$begin:math:display$([0-9a-f:]+)$end:math:display$'){ return $Matches[1] }                    # name [ipv6]
  try{ $uri=$null; if([System.Uri]::TryCreate($t,[System.UriKind]::Absolute,[ref]$uri) -and $uri.Host){ $t=$uri.Host.ToLowerInvariant() } }catch{}
  $t = $t.TrimEnd('.').Trim('[',']')
  $isIPv6=$false; try{ $ip=$null; if([System.Net.IPAddress]::TryParse($t,[ref]$ip)){ $isIPv6=($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) } }catch{}
  if(-not $isIPv6){ if($t -match '^(.+?):(\d+)$'){ $t=$Matches[1] } }
  if($t -match '(^|\s)(\d{1,3}(?:\.\d{1,3}){3})(\s|$)'){ return $Matches[2] }
  return $t
}
function To-DoubleOrNull($v){
  if($v -is [double]){ return [double]$v }
  $s = ('' + $v).Trim()
  if(-not $s){ return $null }
  $d=0.0
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$d)){ return [double]$d }
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::CurrentCulture,[ref]$d)){ return [double]$d }
  return $null
}
function Write-Column2D($ws,[string]$addr,[object[]]$arr){
  [int]$n = if($arr){ [int]$arr.Count } else { 0 }
  if($n -le 0){ return }
  $data = New-Object 'object[,]' ([int]$n),([int]1)
  for($i=0;$i -lt $n;$i++){ $data[$i,0]=$arr[$i] }
  $ws.Range($addr).Resize([int]$n,1).Value2=$data
}
function New-RepeatedArray([object]$value,[int]$count){
  if($count -le 0){ return @() }
  $count=[int]$count
  $a=New-Object object[] $count
  for($i=0;$i -lt $count;$i++){ $a[$i]=$value }
  return $a
}

# ---------- CSV 読み込み ----------
if(-not (Test-Path $CsvPath)){ throw "CSV not found: $CsvPath" }
$data = Import-Csv -Path $CsvPath -Encoding UTF8
if(-not $data -or $data.Count -eq 0){ throw "CSV is empty: $CsvPath" }

# 列解決（表記ゆれ対応）
$headers = @{}
$data[0].PSObject.Properties.Name | ForEach-Object { $headers[$_.ToLowerInvariant()] = $_ }
function Resolve-Col([string[]]$cands){
  foreach($c in $cands){ if($headers.ContainsKey($c)){ return $headers[$c] } }
  foreach($c in $cands){ foreach($k in $headers.Keys){ if($k -like "*$c*"){ return $headers[$k] } } }
  return $null
}

$hn = Resolve-Col @('host','hostname','target','dst_host','dest','remote_host'); if(-not $hn){ throw "host column not found" }
$tn = Resolve-Col @('timestamp','time','datetime','date'); if(-not $tn){ throw "timestamp column not found" }
$in = Resolve-Col @('icmp_avg_ms','rtt_ms','avg_rtt','avg_rtt_ms','icmp_avg','icmp_rtt_ms')
$tc = Resolve-Col @('tcp_ms','tcp_connect_ms','tcp443_ms')
$ht = Resolve-Col @('http_ms','http_head_ms','http_head_rtt_ms')
$dn = Resolve-Col @('dns_ms','dns_lookup_ms','dns_rtt_ms')
$ss = Resolve-Col @('ssid')
$bs = Resolve-Col @('bssid','ap_bssid')
$ap = Resolve-Col @('ap','ap_name','ap_label','ap_hostname')

# ---------- targets.csv 読み込み ----------
function Get-DefaultGatewayIPv4(){
  try{
    $gw = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1
    if($gw){ return $gw.IPv4DefaultGateway.NextHop }
  }catch{}
  return $null
}
function Get-HopN([int]$n){
  try{
    $out = tracert -4 -d -h $n 8.8.8.8 2>$null
    foreach($line in $out){
      if($line -match "^\s*$n\s+\S+\s+\S+\s+\S+\s+(\d{1,3}(?:\.\d{1,3}){3})\s*$"){ return $Matches[1] }
      if($line -match "^\s*$n\s+(\d{1,3}(?:\.\d{1,3}){3})\s*$"){ return $Matches[1] }
    }
  }catch{}
  return $null
}
function Parse-TargetsCsv([string]$path){
  if(-not (Test-Path $path)){ throw "targets.csv not found: $path" }
  $rows = Import-Csv -Path $path -Encoding UTF8
  if(-not $rows -or $rows.Count -eq 0){ throw "targets.csv is empty: $path" }

  $gw = Get-DefaultGatewayIPv4
  $hop2 = Get-HopN 2
  $hop3 = Get-HopN 3

  $list = New-Object System.Collections.Generic.List[object]
  foreach($r in $rows){
    $role  = ('' + $r.role).Trim().ToUpperInvariant()
    $key   = ('' + $r.key ).Trim()
    $label = ('' + $r.label).Trim()
    if(-not $role -or -not $key){ continue }

    if($key -eq '{GATEWAY}' -and $gw){ $key=$gw }
    if($key -eq '{HOP2}'    -and $hop2){ $key=$hop2 }
    if($key -eq '{HOP3}'    -and $hop3){ $key=$hop3 }

    $list.Add([pscustomobject]@{
      Role = $role
      Key  = $key
      KeyNorm = Normalize-Host $key
      Label = (if($label){ $label } else { $key })
    })
  }
  if($list.Count -eq 0){ throw "No valid entries in targets.csv (after placeholders)" }
  return $list
}
$targets = Parse-TargetsCsv $TargetsCsv
# 役割→targets
$roleKeys = @{}
foreach($t in $targets){
  if(-not $roleKeys.ContainsKey($t.Role)){ $roleKeys[$t.Role] = New-Object System.Collections.Generic.List[object] }
  $roleKeys[$t.Role].Add($t)
}

# ---------- フロア判定 ----------
$floorMapHash = @{}
if($FloorMap -and (Test-Path $FloorMap)){
  $fm = Import-Csv -Path $FloorMap -Encoding UTF8
  foreach($r in $fm){
    $b=(''+$r.bssid).ToLowerInvariant()
    $an=(''+$r.ap_name).ToLowerInvariant()
    $f=(''+$r.floor).Trim()
    if($b){ $floorMapHash["bssid::$b"]=$f }
    if($an){ $floorMapHash["ap::$an"]=$f }
  }
}
function Guess-Floor([string]$apName,[string]$bssid,[string]$ssid){
  $key=''
  if($bssid){ $key="bssid::"+$bssid.ToLowerInvariant(); if($floorMapHash.ContainsKey($key)){ return $floorMapHash[$key] } }
  if($apName){ $key="ap::"+$apName.ToLowerInvariant(); if($floorMapHash.ContainsKey($key)){ return $floorMapHash[$key] } }

  $candidates = @($apName,$ssid)
  foreach($c in $candidates){
    $s=(''+$c)
    if([string]::IsNullOrWhiteSpace($s)){ continue }
    $m = [regex]::Match($s,'(?i)\b(\d{1,2})\s*(?:f|階)\b')
    if($m.Success){ return ($m.Groups[1].Value + 'F') }
  }
  return 'Unknown'
}

# ---------- 有効RTT選択（役割別） ----------
function Pick-EffRtt([string]$role, [object]$row){
  $icmp = if($in){ To-DoubleOrNull $row.$in } else { $null }
  $tcp  = if($tc){ To-DoubleOrNull $row.$tc } else { $null }
  $http = if($ht){ To-DoubleOrNull $row.$ht } else { $null }

  if($role -like 'RTR*' -or $role -eq 'L2' -or $role -eq 'L3'){
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
  } else { # SAAS/ZSCALER/その他
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
  }
  return ,@($null,'')
}

# ---------- 集計（Layer用） ----------
if($BucketMinutes -lt 1){ $BucketMinutes = 5 }
[double]$frac = [double]$BucketMinutes / 1440.0
$ciCur=[System.Globalization.CultureInfo]::CurrentCulture
$ciInv=[System.Globalization.CultureInfo]::InvariantCulture

$roleOrder = @('L2','L3','RTR_LAN','RTR_WAN','ZSCALER','SAAS') | Where-Object { $roleKeys.ContainsKey($_) }

$buckets = @{}  # bucket -> role -> List<double>

foreach($row in $data){
  $hraw = '' + $row.$hn; if(-not $hraw){ continue }
  $hnorm = Normalize-Host $hraw
  $ts = '' + $row.$tn; if(-not $ts){ continue }
  try{ $dt=[datetime]::Parse($ts,$ciCur) }catch{ try{ $dt=[datetime]::Parse($ts,$ciInv) }catch{ continue } }
  [double]$tOa = $dt.ToOADate()
  [double]$bucket = [math]::Floor($tOa / $frac) * $frac

  foreach($role in $roleOrder){
    $matched=$false
    foreach($t in $roleKeys[$role]){
      if($hnorm -eq $t.KeyNorm -or $hnorm.Contains($t.KeyNorm) -or $t.KeyNorm.Contains($hnorm)){ $matched=$true; break }
    }
    if(-not $matched){ continue }
    $pair = Pick-EffRtt $role $row
    $val = $pair[0]; if($val -eq $null){ continue }
    if(-not $buckets.ContainsKey($bucket)){ $buckets[$bucket]=@{} }
    if(-not $buckets[$bucket].ContainsKey($role)){ $buckets[$bucket][$role]=New-Object System.Collections.Generic.List[double] }
    $buckets[$bucket][$role].Add([double]$val)
  }
}
if($buckets.Count -eq 0){ Write-Warning "No layer data matched targets.csv"; }

# 平均系列とΔ
$X=@(); $series=@{}; foreach($r in $roleOrder){ $series[$r]=@() }
$DeltaNames=@('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD'); $delta=@{}; foreach($d in $DeltaNames){ $delta[$d]=@() }
$keys=@($buckets.Keys) | Sort-Object {[double]$_}
foreach($k in $keys){
  $X += [double]$k
  $valsThis=@{}
  foreach($r in $roleOrder){
    if($buckets[$k].ContainsKey($r)){
      $arr = $buckets[$k][$r].ToArray()
      [double]$avg = ($arr | Measure-Object -Average).Average
      $series[$r] += [double]$avg
      $valsThis[$r] = [double]$avg
    } else {
      $series[$r] += $null
      $valsThis[$r] = $null
    }
  }
  function SubOrNull($a,$b){ if($a -ne $null -and $b -ne $null){ return [double]($a-$b) } else { return $null } }
  $l2  = if($valsThis.ContainsKey('L2')){ $valsThis['L2'] }else{$null}
  $l3  = if($valsThis.ContainsKey('L3')){ $valsThis['L3'] }else{$null}
  $lan = if($valsThis.ContainsKey('RTR_LAN')){ $valsThis['RTR_LAN'] }else{$null}
  $wan = if($valsThis.ContainsKey('RTR_WAN')){ $valsThis['RTR_WAN'] }else{$null}
  $saas= if($valsThis.ContainsKey('SAAS')){ $valsThis['SAAS'] }else{$null}
  $delta['DELTA_L3']      += (SubOrNull $l3  $l2)
  $delta['DELTA_RTR_LAN'] += (SubOrNull $lan $l3)
  $delta['DELTA_RTR_WAN'] += (SubOrNull $wan $lan)
  $delta['DELTA_CLOUD']   += (SubOrNull $saas $wan)
}

# ---------- Excel 出力 ----------
[int]$xlXYScatterLines=74; [int]$xlLegendBottom=-4107; [int]$xlCategory=1; [int]$xlValue=2; [int]$msoLineDash=4
$excel=$null; $wb=$null
try{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible=[bool]$Visible
  $excel.DisplayAlerts=$false
  $wb = $excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }

  # ---- LayerSeries ----
  if($X.Count -gt 0){
    $ws1 = $wb.Worksheets.Item(1); $ws1.Name = 'LayerSeries'
    $n=$X.Count
    $ws1.Cells(1,1).Value2='timestamp'
    $arrX = New-Object 'object[,]' $n,1
    for($i=0;$i -lt $n;$i++){ $arrX[$i,0]=$X[$i] }
    $ws1.Range('A2').Resize($n,1).Value2=$arrX
    $ws1.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm'
    $col=2
    foreach($r in $roleOrder){
      $ws1.Cells(1,$col).Value2=$r
      $vals=$series[$r]
      $arr = New-Object 'object[,]' $n,1
      for($i=0;$i -lt $n;$i++){ $arr[$i,0]=$vals[$i] }
      $ws1.Range($ws1.Cells(2,$col),$ws1.Cells(1+$n,$col)).Value2=$arr
      $col++
    }
    $ws1.Columns.AutoFit() | Out-Null
    $ch1 = $ws1.ChartObjects().Add(320,10,900,330)
    $c1=$ch1.Chart; $c1.ChartType=$xlXYScatterLines; $c1.HasTitle=$true; $c1.ChartTitle.Text='Layer RTT (bucket avg)'; $c1.Legend.Position=$xlLegendBottom
    try{ $c1.SeriesCollection().Delete() }catch{}
    $endRow=1+$n; $sCol=2
    foreach($r in $roleOrder){
      $s=$c1.SeriesCollection().NewSeries()
      $s.Name=$r
      $s.XValues=$ws1.Range(("A2:A{0}" -f $endRow))
      $s.Values =$ws1.Range(($ws1.Cells(2,$sCol).Address()+":"+$ws1.Cells($endRow,$sCol).Address()))
      $sCol++
    }
    try{
      $v=$c1.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
      $x=$c1.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}
    # ---- DeltaSeries ----
    $ws2 = $wb.Worksheets.Add(); $ws2.Name='DeltaSeries'
    $ws2.Cells(1,1).Value2='timestamp'
    $ws2.Range('A2').Resize($n,1).Value2=$arrX
    $ws2.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm'
    $col=2
    foreach($dName in @('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD')){
      $ws2.Cells(1,$col).Value2=$dName
      $vals=$delta[$dName]
      $arr = New-Object 'object[,]' $n,1
      for($i=0;$i -lt $n;$i++){ $arr[$i,0]=$vals[$i] }
      $ws2.Range($ws2.Cells(2,$col),$ws2.Cells(1+$n,$col)).Value2=$arr
      $col++
    }
    $ws2.Columns.AutoFit() | Out-Null
    $ch2 = $ws2.ChartObjects().Add(320,10,900,330)
    $c2=$ch2.Chart; $c2.ChartType=$xlXYScatterLines; $c2.HasTitle=$true; $c2.ChartTitle.Text='Layer Δ (bucket avg)'; $c2.Legend.Position=$xlLegendBottom
    try{ $c2.SeriesCollection().Delete() }catch{}
    $endRow=1+$n; $sCol=2
    foreach($dName in @('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD')){
      $s=$c2.SeriesCollection().NewSeries()
      $s.Name=$dName
      $s.XValues=$ws2.Range(("A2:A{0}" -f $endRow))
      $s.Values =$ws2.Range(($ws2.Cells(2,$sCol).Address()+":"+$ws2.Cells($endRow,$sCol).Address()))
      $sCol++
    }
    try{
      $v=$c2.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
      $x=$c2.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}
  } else {
    # Xがない場合でも後続の HostSheets は作る
    $wb.Worksheets.Item(1).Name = 'INDEX'
  }

  # ---- HostSheets（ターゲットごと & フロア別系列）----
  function Ensure-UniqueSheetName($wb,[string]$base){
    $sn = Sanitize-SheetName $base
    $orig=$sn; $i=2
    while($true){
      try{ $null=$wb.Worksheets.Item($sn); $sn = Sanitize-SheetName ($orig+"_"+$i); $i++ }catch{ break }
    }
    return $sn
  }

  $created=@()
  foreach($t in $targets){
    # このターゲットにマッチする行を抽出
    $rows = @()
    foreach($row in $data){
      $hraw = '' + $row.$hn; if(-not $hraw){ continue }
      $hnorm = Normalize-Host $hraw
      if($hnorm -eq $t.KeyNorm -or $hnorm.Contains($t.KeyNorm) -or $t.KeyNorm.Contains($hnorm)){
        $rows += ,$row
      }
    }
    if($rows.Count -eq 0){ continue }

    # フロア別に X/Y を構築（役割に応じ有効RTT採用）
    $byFloor = @{} # floor -> @{ times = List<double>; vals = List<double> }
    foreach($row in $rows){
      $ts = '' + $row.$tn; if(-not $ts){ continue }
      try{ $dt=[datetime]::Parse($ts,$ciCur) }catch{ try{ $dt=[datetime]::Parse($ts,$ciInv) }catch{ continue } }
      [double]$x = $dt.ToOADate()
      $pair = Pick-EffRtt $t.Role $row
      $y = $pair[0]; if($y -eq $null){ continue }

      $apName = if($ap){ ''+$row.$ap } else { '' }
      $bssid  = if($bs){ ''+$row.$bs } else { '' }
      $ssidv  = if($ss){ ''+$row.$ss } else { '' }
      $floor = Guess-Floor $apName $bssid $ssidv

      if(-not $byFloor.ContainsKey($floor)){
        $byFloor[$floor]=@{ times=New-Object System.Collections.Generic.List[double]; vals=New-Object System.Collections.Generic.List[double] }
      }
      $byFloor[$floor].times.Add([double]$x)
      $byFloor[$floor].vals.Add([double]$y)
    }

    if($byFloor.Keys.Count -eq 0){ continue }

    $sn = Ensure-UniqueSheetName $wb $t.Label
    $ws = $wb.Worksheets.Add()
    $ws.Name = $sn

    # カラム配置: floorごとに 2列(pair) → time_X, rtt_X ... 最後に time_all, threshold
    $col = 1
    $floorList = @($byFloor.Keys) | Sort-Object
    $unionTimes = New-Object System.Collections.Generic.List[double]
    foreach($fl in $floorList){
      $times = @($byFloor[$fl].times.ToArray()) | Sort-Object
      $vals  = @($byFloor[$fl].vals.ToArray())  | Sort-Object @{Expression={$times.IndexOf($_)}} # 安全側: 同じ順で

      # 再ソート: times と vals を同じ順に（簡易にpair再構築）
      $pairs=@()
      for($i=0;$i -lt $times.Count;$i++){ $pairs += [pscustomobject]@{ t=[double]$byFloor[$fl].times[$i]; r=[double]$byFloor[$fl].vals[$i] } }
      $pairs = $pairs | Sort-Object t
      $times = @(); $vals=@()
      foreach($p in $pairs){ $times += [double]$p.t; $vals += [double]$p.r }

      $ws.Cells(1,$col).Value2=("time_{0}" -f $fl); $ws.Cells(1,$col+1).Value2=("rtt_{0}" -f $fl)
      Write-Column2D $ws ($ws.Cells(2,$col).Address())  $times
      Write-Column2D $ws ($ws.Cells(2,$col+1).Address()) $vals
      try{ $ws.Range($ws.Cells(2,$col),$ws.Cells(1+$times.Count,$col)).NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
      try{ $ws.Range($ws.Cells(2,$col+1),$ws.Cells(1+$times.Count,$col+1)).NumberFormatLocal='0.0' }catch{}
      $col += 2

      foreach($x in $times){ $unionTimes.Add([double]$x) }
    }
    $ws.Columns.AutoFit() | Out-Null

    # 閾値列（unionTimes で）
    $unionSorted = @($unionTimes.ToArray()) | Sort-Object
    $ws.Cells(1,$col).Value2='time_all'; $ws.Cells(1,$col+1).Value2='threshold_ms'
    Write-Column2D $ws ($ws.Cells(2,$col).Address()) $unionSorted
    Write-Column2D $ws ($ws.Cells(2,$col+1).Address()) (New-RepeatedArray -value ([double]$ThresholdMs) -count $unionSorted.Count)
    try{ $ws.Range($ws.Cells(2,$col),$ws.Cells(1+$unionSorted.Count,$col)).NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
    try{ $ws.Range($ws.Cells(2,$col+1),$ws.Cells(1+$unionSorted.Count,$col+1)).NumberFormatLocal='0.0' }catch{}

    # グラフ
    $ch=$ws.ChartObjects().Add(320,10,900,330)
    $c=$ch.Chart; $c.ChartType=$xlXYScatterLines; $c.HasTitle=$true
    $c.ChartTitle.Text = ("{0} - RTT by Floor" -f $t.Label)
    $c.Legend.Position=$xlLegendBottom
    try{ $c.SeriesCollection().Delete() }catch{}

    # フロア系列
    $colIdx=1
    foreach($fl in $floorList){
      $s=$c.SeriesCollection().NewSeries()
      $s.Name=("RTT ({0})" -f $fl)
      # X: colIdx, Y: colIdx+1
      $endRowTime = ($ws.Cells($ws.Rows.Count,$colIdx).End(-4162)).Row # xlUp=-4162
      $endRowVal  = ($ws.Cells($ws.Rows.Count,$colIdx+1).End(-4162)).Row
      $endRow = [Math]::Max($endRowTime,$endRowVal)
      $s.XValues=$ws.Range(($ws.Cells(2,$colIdx).Address()+":"+$ws.Cells($endRow,$colIdx).Address()))
      $s.Values =$ws.Range(($ws.Cells(2,$colIdx+1).Address()+":"+$ws.Cells($endRow,$colIdx+1).Address()))
      $colIdx+=2
    }
    # 閾値系列（最後の2列）
    $s2=$c.SeriesCollection().NewSeries()
    $s2.Name=("threshold {0} ms" -f [int]$ThresholdMs)
    $endRow = ($ws.Cells($ws.Rows.Count,$col).End(-4162)).Row
    $s2.XValues=$ws.Range(($ws.Cells(2,$col).Address()+":"+$ws.Cells($endRow,$col).Address()))
    $s2.Values =$ws.Range(($ws.Cells(2,$col+1).Address()+":"+$ws.Cells($endRow,$col+1).Address()))
    try{ $s2.Format.Line.ForeColor.RGB=255; $s2.Format.Line.Weight=1.5; $s2.Format.Line.DashStyle=$msoLineDash }catch{}

    try{
      $v=$c.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
      $x=$c.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}

    $created += $sn
  }

  # ---- INDEX シート ----
  $wsIdx=$wb.Worksheets.Add(); $wsIdx.Name='INDEX'
  $wsIdx.Cells(1,1).Value2='Sheets'
  $r=2
  foreach($w in @('LayerSeries','DeltaSeries')){
    try{ $null=$wb.Worksheets.Item($w); $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$w'!A1",'',$w) | Out-Null; $r++ }catch{}
  }
  foreach($sn in $created){
    $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$sn'!A1",'',$sn) | Out-Null; $r++
  }

  $wb.SaveAs($Output)
  Write-Host "Output: $Output"
}
catch{
  Write-Error $_
  throw
}
finally{
  if($wb){ try{ $wb.Close($false) }catch{}; Release-Com $wb; $wb=$null }
  if($excel){ try{ $excel.Quit() }catch{}; Release-Com $excel; $excel=$null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}