<#
Generate-TeamsNet-Report.ps1  (PowerShell 5.1 互換・Excel COMは必ず解放・メモリ対策/COM安全化版)

機能:
- teams_net_quality.csv を集計し、同一Excelブックに以下を出力
  1) LayerSeries: L2/L3/RTR_LAN/RTR_WAN/ZSCALER/SAAS の時間バケット平均
  2) DeltaSeries: 上記の区間差(Δ) = L3-L2, RTR_LAN-L3, RTR_WAN-RTR_LAN, SAAS-RTR_WAN
  3) ホスト別シート: targets.csv の各行(=ターゲット)ごとに1シート、AP名/BSSID/SSIDからフロアを推定 or floors.csv で色分けし、時系列グラフ＋しきい値線を描画
- SAAS/Zscaler は ICMPが得られない想定のため TCP/HTTP を優先、L2/L3/RTR* は ICMP を優先
- 「if を式として使う」書き方は不使用（PS5.1 準拠）
- Excel への書き込みは**分割（チャンク）**で実施し、**軽量モード**で描画＆再計算を停止（メモリ不足対策）
- Write-Column2D が **COM Range / 配列 / 文字列 / 単一値**を自動正規化（System.__ComObject 受け取り時の引数変換エラー回避）

使い方(例):
  powershell -NoProfile -ExecutionPolicy Bypass `
    -File .\Generate-TeamsNet-Report.ps1 `
    -CsvPath "$Env:LOCALAPPDATA\TeamsNet\teams_net_quality.csv" `
    -TargetsCsv "..\targets.csv" `
    -FloorMap "..\floors.csv" `
    -Output ".\Output\TeamsNet-Report.xlsx" `
    -BucketMinutes 5 `
    -ThresholdMs 100
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$CsvPath,
  [Parameter(Mandatory=$true)][string]$TargetsCsv,
  [Parameter(Mandatory=$true)][string]$Output,
  [int]$BucketMinutes = 5,
  [int]$ThresholdMs = 100,
  [string]$FloorMap,
  [switch]$Visible
)

# ---------------- 共通設定 ----------------
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

function Release-Com([object]$obj){
  if($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)){
    try{ [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }catch{}
  }
}

function Set-ExcelPerfMode($app, [bool]$on){
  try{
    if($on){
      $app.ScreenUpdating = $false
      $app.DisplayStatusBar = $false
      $app.EnableEvents = $false
      $app.Calculation = -4135   # xlCalculationManual
    } else {
      $app.Calculation = -4105   # xlCalculationAutomatic
      $app.EnableEvents = $true
      $app.ScreenUpdating = $true
      $app.DisplayStatusBar = $true
    }
  }catch{}
}

function Sanitize-SheetName([string]$name){
  if(-not $name){ return 'Sheet' }
  $n = $name -replace '[:\\/\?\*$begin:math:display$$end:math:display$]','_'
  if($n.Length -gt 31){ $n=$n.Substring(0,31) }
  if($n -match '^\s*$'){ $n='Sheet' }
  return $n
}

function Normalize-Host([string]$s){
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  $t = $s.Trim().Trim('"',"'").ToLowerInvariant()
  if ([string]::IsNullOrEmpty($t)) { return '' }

  if ($t -match '$begin:math:text$([0-9]{1,3}(?:\\.[0-9]{1,3}){3})$end:math:text$') { return $Matches[1] } # name (ipv4)
  if ($t -match '$begin:math:display$([0-9a-f:]+)$end:math:display$')                   { return $Matches[1] } # name [ipv6]

  try {
    $uri = $null
    if ([System.Uri]::TryCreate($t, [System.UriKind]::Absolute, [ref]$uri) -and $uri.Host) {
      $t = $uri.Host.ToLowerInvariant()
    }
  } catch {}

  $t = $t.TrimEnd('.').Trim('[',']')
  if ([string]::IsNullOrEmpty($t)) { return '' }

  try {
    $ip = $null
    if ([System.Net.IPAddress]::TryParse($t, [ref]$ip)) { return $t }
  } catch {}

  if ($t -match '^(.+?):(\d+)$' -and $t -notmatch '^$begin:math:display$.+$end:math:display$') { $t = $Matches[1] }
  if ($t -match '(^|\s|$begin:math:text$)(\\d{1,3}(?:\\.\\d{1,3}){3})(\\s|$end:math:text$|:|$)') { return $Matches[2] }

  return $t
}

function To-DoubleOrNull($v){
  if($v -is [double]){ return [double]$v }
  $s=(''+$v).Trim()
  if(-not $s){ return $null }
  $d=0.0
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$d)){ return [double]$d }
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::CurrentCulture,[ref]$d)){ return [double]$d }
  return $null
}

# 分割書き込み（メモリ節約＋COM安全化）
function Write-Column2D($ws,[string]$addr,$seq,[int]$ChunkSize=20000){
  if ($null -eq $seq) { return }

  # --- $seq を 1 次元 List<object> に正規化 ---
  $list = New-Object System.Collections.Generic.List[object]

  # COM Range / Variant を受け取った場合
  if ([System.Runtime.InteropServices.Marshal]::IsComObject($seq)) {
    try {
      $v = $seq.Value2
      if ($null -eq $v) { return }
      if ($v -is [object[,]]) {
        $rows=$v.GetLength(0); $cols=$v.GetLength(1)
        # 1列目を使う（複数列時）
        for($r=1;$r -le $rows;$r++){ $list.Add($v[$r,1]) }
      } else {
        $list.Add($v)
      }
    } catch {
      # 最低限のフォールバック
      $list.Add((''+$seq))
    }
  }
  elseif ($seq -is [System.Collections.IEnumerable] -and -not ($seq -is [string])) {
    foreach($e in $seq){ $list.Add($e) }
  }
  else {
    $list.Add($seq)
  }

  $n = $list.Count
  if ($n -le 0) { return }

  $start = $ws.Range($addr)
  $idx=0
  while($idx -lt $n){
    $take=[Math]::Min($ChunkSize,$n-$idx)
    $block = New-Object 'object[,]' $take, 1
    for($r=0;$r -lt $take;$r++){ $block[$r,0]=$list[$idx+$r] }
    $start.Resize($take,1).Value2 = $block
    $start = $start.Offset($take, 0)
    $idx += $take
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function New-RepeatedArray([object]$value,[int]$count){
  if($count -le 0){ return @() }
  $a = New-Object object[] $count
  for($i=0;$i -lt $count;$i++){ $a[$i]=$value }
  return $a
}

function Format-Err([System.Management.Automation.ErrorRecord]$e){
  $ii=$e.InvocationInfo
  $parts=@("[ERROR] $($e.FullyQualifiedErrorId)")
  if($ii){ $parts += " at line $($ii.ScriptLineNumber) char $($ii.OffsetInLine): $($ii.Line)" }
  $ex=$e.Exception
  if($ex){ $parts += "$($ex.GetType().FullName): $($ex.Message)" }
  return ($parts -join "`r`n")
}

function Sub-OrNull($a,$b){
  if($a -ne $null -and $b -ne $null){ return [double]($a-$b) }
  return $null
}

# Hashtable/OrderedDictionary/Dictionary などで安全にキー存在確認
function Test-MapHasKey($map, $key){
  if ($null -eq $map) { return $false }
  try {
    $methods = $map.PSObject.Methods.Name
  } catch {
    try { return ($map.Keys -contains $key) } catch { return $false }
  }
  if ($methods -contains 'ContainsKey') { return [bool]$map.ContainsKey($key) }
  if ($methods -contains 'Contains')    { return [bool]$map.Contains($key) }
  try { return ($map.Keys -contains $key) } catch { return $false }
}

# ---------------- CSV 読み込み ----------------
if(-not (Test-Path $CsvPath)){ throw "CSV not found: $CsvPath" }
$data = Import-Csv -Path $CsvPath -Encoding UTF8
if(-not $data -or $data.Count -eq 0){ throw "CSV is empty: $CsvPath" }

# 列解決（表記ゆれ対応）
$headers=@{}; $data[0].PSObject.Properties.Name | ForEach-Object { $headers[$_.ToLowerInvariant()] = $_ }
function Resolve-Col([string[]]$cands){
  foreach($c in $cands){
    if(Test-MapHasKey $headers $c){ return $headers[$c] }
  }
  foreach($c in $cands){
    foreach($k in $headers.Keys){
      if($k -like "*$c*"){ return $headers[$k] }
    }
  }
  return $null
}
$colHost = Resolve-Col @('host','hostname','target','dst_host','dest','remote_host'); if(-not $colHost){ throw "host column not found" }
$colTime = Resolve-Col @('timestamp','time','datetime','date'); if(-not $colTime){ throw "timestamp column not found" }
$colIcmp = Resolve-Col @('icmp_avg_ms','rtt_ms','avg_rtt','avg_rtt_ms','icmp_avg','icmp_rtt_ms')
$colTcp  = Resolve-Col @('tcp_ms','tcp_connect_ms','tcp443_ms')
$colHttp = Resolve-Col @('http_ms','http_head_ms','http_head_rtt_ms')
$colDns  = Resolve-Col @('dns_ms','dns_lookup_ms','dns_rtt_ms')
$colSsid = Resolve-Col @('ssid')
$colBssid= Resolve-Col @('bssid','ap_bssid')
$colAp   = Resolve-Col @('ap','ap_name','ap_label','ap_hostname')

# ---------------- targets.csv 読み込み ----------------
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
  $gw  = Get-DefaultGatewayIPv4
  $hop2= Get-HopN 2
  $hop3= Get-HopN 3
  $list = New-Object System.Collections.Generic.List[object]
  foreach($r in $rows){
    $role  = (''+$r.role).Trim().ToUpperInvariant()
    $key   = (''+$r.key ).Trim()
    $label = (''+$r.label).Trim()
    if(-not $role -or -not $key){ continue }
    if($key -eq '{GATEWAY}' -and $gw){ $key=$gw }
    if($key -eq '{HOP2}'    -and $hop2){ $key=$hop2 }
    if($key -eq '{HOP3}'    -and $hop3){ $key=$hop3 }

    $lab = $key
    if (-not [string]::IsNullOrWhiteSpace($label)) { $lab = $label }

    $list.Add([pscustomobject]@{
      Role   = $role
      Key    = $key
      KeyNorm= (Normalize-Host $key)
      Label  = $lab
    })
  }
  if($list.Count -eq 0){ throw "No valid entries in targets.csv (after placeholders)" }
  return $list
}
$targets = Parse-TargetsCsv $TargetsCsv

# 役割→targets
$roleKeys=@{}
foreach($t in $targets){
  if(-not (Test-MapHasKey $roleKeys $t.Role)){ $roleKeys[$t.Role] = New-Object System.Collections.Generic.List[object] }
  $roleKeys[$t.Role].Add($t)
}

# ---------------- フロア判定 ----------------
$floorMap=@{}
if($FloorMap -and (Test-Path $FloorMap)){
  $fm = Import-Csv -Path $FloorMap -Encoding UTF8
  foreach($r in $fm){
    $b=(''+$r.bssid).ToLowerInvariant()
    $an=(''+$r.ap_name).ToLowerInvariant()
    $f=(''+$r.floor).Trim()
    if($b){ $floorMap["bssid::$b"]=$f }
    if($an){ $floorMap["ap::$an"]=$f }
  }
}
function Guess-Floor([string]$apName,[string]$bssid,[string]$ssid){
  if($bssid){
    $k="bssid::"+$bssid.ToLowerInvariant()
    if(Test-MapHasKey $floorMap $k){ return $floorMap[$k] }
  }
  if($apName){
    $k="ap::"+$apName.ToLowerInvariant()
    if(Test-MapHasKey $floorMap $k){ return $floorMap[$k] }
  }
  foreach($c in @($apName,$ssid)){
    $s=(''+$c)
    if([string]::IsNullOrWhiteSpace($s)){ continue }
    $m=[regex]::Match($s,'(?i)\b(\d{1,2})\s*(?:f|階)\b')
    if($m.Success){ return ($m.Groups[1].Value + 'F') }
  }
  return 'Unknown'
}

# ---------------- 有効RTT 選択（役割別） ----------------
function Pick-EffRtt([string]$role,[object]$row){
  $icmp = $null; if ($colIcmp) { $icmp = To-DoubleOrNull $row.$colIcmp }
  $tcp  = $null; if ($colTcp ) { $tcp  = To-DoubleOrNull $row.$colTcp  }
  $http = $null; if ($colHttp) { $http = To-DoubleOrNull $row.$colHttp }

  if($role -like 'RTR*' -or $role -eq 'L2' -or $role -eq 'L3'){
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
  } else {
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
  }
  return ,@($null,'')
}

# ---------------- 集計（Layer 用） ----------------
if($BucketMinutes -lt 1){ $BucketMinutes=5 }
[double]$frac = [double]$BucketMinutes / 1440.0
$ciCur=[System.Globalization.CultureInfo]::CurrentCulture
$ciInv=[System.Globalization.CultureInfo]::InvariantCulture

# 存在する役割のみを順序付きで採用
$roleOrderAll = @('L2','L3','RTR_LAN','RTR_WAN','ZSCALER','SAAS')
$roleOrder=@()
foreach($r in $roleOrderAll){ if(Test-MapHasKey $roleKeys $r){ $roleOrder += $r } }

# bucket(double OAdate) -> role -> List<double>
$buckets=@{}
$X=@()

foreach($row in $data){
  $hraw = ''+$row.$colHost; if(-not $hraw){ continue }
  $hnorm = Normalize-Host $hraw
  $ts = ''+$row.$colTime; if(-not $ts){ continue }
  try{ $dt=[datetime]::Parse($ts,$ciCur) }catch{ try{ $dt=[datetime]::Parse($ts,$ciInv) }catch{ continue } }
  [double]$tOa = $dt.ToOADate()
  [double]$bucket = [math]::Floor($tOa / $frac) * $frac

  foreach($r in $roleOrder){
    $matched=$false
    foreach($t in $roleKeys[$r]){
      if($hnorm -eq $t.KeyNorm -or $hnorm.Contains($t.KeyNorm) -or $t.KeyNorm.Contains($hnorm)){ $matched=$true; break }
    }
    if(-not $matched){ continue }
    $pair = Pick-EffRtt $r $row
    $val = $pair[0]; if($val -eq $null){ continue }
    if(-not (Test-MapHasKey $buckets $bucket)){ $buckets[$bucket]=@{} }
    if(-not (Test-MapHasKey $buckets[$bucket] $r)){ $buckets[$bucket][$r]=New-Object System.Collections.Generic.List[double] }
    $buckets[$bucket][$r].Add([double]$val)
  }
}

# 平均＆Δ系列
$series=@{}; foreach($r in $roleOrder){ $series[$r]=@() }
$DeltaNames=@('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD'); $delta=@{}; foreach($d in $DeltaNames){ $delta[$d]=@() }
$keys=@($buckets.Keys) | Sort-Object {[double]$_}
foreach($k in $keys){
  $X += [double]$k
  $valsThis=@{}
  foreach($r in $roleOrder){
    if(Test-MapHasKey $buckets[$k] $r){
      $arr=$buckets[$k][$r].ToArray()
      [double]$avg = ($arr | Measure-Object -Average).Average
      $series[$r] += [double]$avg
      $valsThis[$r]=[double]$avg
    }else{
      $series[$r] += $null
      $valsThis[$r]=$null
    }
  }
  $l2  = $null; if(Test-MapHasKey $valsThis 'L2')      { $l2  = $valsThis['L2'] }
  $l3  = $null; if(Test-MapHasKey $valsThis 'L3')      { $l3  = $valsThis['L3'] }
  $lan = $null; if(Test-MapHasKey $valsThis 'RTR_LAN') { $lan = $valsThis['RTR_LAN'] }
  $wan = $null; if(Test-MapHasKey $valsThis 'RTR_WAN') { $wan = $valsThis['RTR_WAN'] }
  $saas= $null; if(Test-MapHasKey $valsThis 'SAAS')    { $saas= $valsThis['SAAS'] }

  $delta['DELTA_L3']      += (Sub-OrNull $l3  $l2)
  $delta['DELTA_RTR_LAN'] += (Sub-OrNull $lan $l3)
  $delta['DELTA_RTR_WAN'] += (Sub-OrNull $wan $lan)
  $delta['DELTA_CLOUD']   += (Sub-OrNull $saas $wan)
}

# ---------------- Excel 出力 ----------------
[int]$xlXYScatterLines=74
[int]$xlLegendBottom=-4107
[int]$xlCategory=1
[int]$xlValue=2
[int]$msoLineDash=4

$excel=$null; $wb=$null
try{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible=[bool]$Visible
  $excel.DisplayAlerts=$false
  Set-ExcelPerfMode $excel $true

  $wb = $excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }

  # ---- LayerSeries ----
  if($X.Count -gt 0){
    $ws1=$wb.Worksheets.Item(1); $ws1.Name='LayerSeries'
    $n=$X.Count
    $ws1.Cells(1,1).Value2='timestamp'
    Write-Column2D $ws1 'A2' $X
    $ws1.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='mm/dd hh:mm'
    $col=2
    foreach($r in $roleOrder){
      $ws1.Cells(1,$col).Value2=$r
      $vals=$series[$r]
      Write-Column2D $ws1 ($ws1.Cells(2,$col).Address()) $vals
      $col++
    }
    $ws1.Columns.AutoFit() | Out-Null
    $ch1=$ws1.ChartObjects().Add(320,10,900,330); $c1=$ch1.Chart
    $c1.ChartType=$xlXYScatterLines; $c1.HasTitle=$true; $c1.ChartTitle.Text='Layer RTT (bucket avg)'; $c1.Legend.Position=$xlLegendBottom
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
    $ws2=$wb.Worksheets.Add(); $ws2.Name='DeltaSeries'
    $ws2.Cells(1,1).Value2='timestamp'
    Write-Column2D $ws2 'A2' $X
    $ws2.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='mm/dd hh:mm'
    $col=2
    foreach($d in @('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD')){
      $ws2.Cells(1,$col).Value2=$d
      $vals=$delta[$d]
      Write-Column2D $ws2 ($ws2.Cells(2,$col).Address()) $vals
      $col++
    }
    $ws2.Columns.AutoFit() | Out-Null
    $ch2=$ws2.ChartObjects().Add(320,10,900,330); $c2=$ch2.Chart
    $c2.ChartType=$xlXYScatterLines; $c2.HasTitle=$true; $c2.ChartTitle.Text='Layer Δ (bucket avg)'; $c2.Legend.Position=$xlLegendBottom
    try{ $c2.SeriesCollection().Delete() }catch{}
    $endRow=1+$n; $sCol=2
    foreach($d in @('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD')){
      $s=$c2.SeriesCollection().NewSeries()
      $s.Name=$d
      $s.XValues=$ws2.Range(("A2:A{0}" -f $endRow))
      $s.Values =$ws2.Range(($ws2.Cells(2,$sCol).Address()+":"+$ws2.Cells($endRow,$sCol).Address()))
      $sCol++
    }
    try{
      $v=$c2.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
      $x=$c2.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}
  } else {
    $wb.Worksheets.Item(1).Name='INDEX'
  }

  # ---- ホスト別シート（フロア色分け） ----
  function Ensure-UniqueSheetName($wb,[string]$base){
    $sn=Sanitize-SheetName $base
    $orig=$sn; $i=2
    while($true){
      $exists=$false
      for($ix=1;$ix -le $wb.Worksheets.Count;$ix++){
        if($wb.Worksheets.Item($ix).Name -eq $sn){ $exists=$true; break }
      }
      if(-not $exists){ break }
      $sn=Sanitize-SheetName ($orig+"_"+$i); $i++
    }
    return $sn
  }

  $created=@()
  foreach($t in $targets){
    # 該当行抽出
    $rows=@()
    foreach($row in $data){
      $hraw=''+$row.$colHost; if(-not $hraw){ continue }
      $hnorm=Normalize-Host $hraw
      if($hnorm -eq $t.KeyNorm -or $hnorm.Contains($t.KeyNorm) -or $t.KeyNorm.Contains($hnorm)){
        $rows+=,$row
      }
    }
    if($rows.Count -eq 0){ continue }

    # floor -> { times(List<double>), vals(List<double>) }
    $byFloor=@{}
    foreach($row in $rows){
      $ts=''+$row.$colTime; if(-not $ts){ continue }
      try{ $dt=[datetime]::Parse($ts,$ciCur) }catch{ try{ $dt=[datetime]::Parse($ts,$ciInv) }catch{ continue } }
      [double]$x=$dt.ToOADate()

      $pair=Pick-EffRtt $t.Role $row
      $y=$pair[0]; if($y -eq $null){ continue }

      $apName = ''; if ($colAp)    { $apName = ''+$row.$colAp }
      $bssid  = ''; if ($colBssid) { $bssid  = ''+$row.$colBssid }
      $ssidv  = ''; if ($colSsid)  { $ssidv  = ''+$row.$colSsid }
      $floor  = Guess-Floor $apName $bssid $ssidv

      if(-not (Test-MapHasKey $byFloor $floor)){
        $byFloor[$floor]=@{ times=New-Object System.Collections.Generic.List[double]; vals=New-Object System.Collections.Generic.List[double] }
      }
      $byFloor[$floor].times.Add([double]$x)
      $byFloor[$floor].vals.Add([double]$y)
    }
    if($byFloor.Keys.Count -eq 0){ continue }

    $ws=$wb.Worksheets.Add()
    $ws.Name = Ensure-UniqueSheetName $wb $t.Label

    # 各フロア列を書き出し（time_X, rtt_X）
    $col=1
    $floorList=@($byFloor.Keys) | Sort-Object
    $unionTimes = New-Object System.Collections.Generic.List[double]
    foreach($fl in $floorList){
      # ペア化して時刻で整列（X/Y同期）
      $pairs=@()
      for($i=0;$i -lt $byFloor[$fl].times.Count;$i++){
        $pairs += [pscustomobject]@{ t=[double]$byFloor[$fl].times[$i]; r=[double]$byFloor[$fl].vals[$i] }
      }
      $pairs = $pairs | Sort-Object t
      $times=@(); $vals=@()
      foreach($p in $pairs){ $times += [double]$p.t; $vals += [double]$p.r }

      $ws.Cells(1,$col).Value2=("time_{0}" -f $fl)
      $ws.Cells(1,$col+1).Value2=("rtt_{0}" -f $fl)
      Write-Column2D $ws ($ws.Cells(2,$col).Address())  $times
      Write-Column2D $ws ($ws.Cells(2,$col+1).Address()) $vals
      try{ $ws.Range($ws.Cells(2,$col),$ws.Cells(1+$times.Count,$col)).NumberFormatLocal='mm/dd hh:mm' }catch{}
      try{ $ws.Range($ws.Cells(2,$col+1),$ws.Cells(1+$times.Count,$col+1)).NumberFormatLocal='0.0' }catch{}
      foreach($xv in $times){ $unionTimes.Add([double]$xv) }
      $col+=2
    }
    $ws.Columns.AutoFit() | Out-Null

    # しきい値シリーズ
    $unionSorted=@($unionTimes.ToArray()) | Sort-Object
    $ws.Cells(1,$col).Value2='time_all'
    $ws.Cells(1,$col+1).Value2='threshold_ms'
    Write-Column2D $ws ($ws.Cells(2,$col).Address()) $unionSorted
    Write-Column2D $ws ($ws.Cells(2,$col+1).Address()) (New-RepeatedArray -value ([double]$ThresholdMs) -count $unionSorted.Count)
    try{ $ws.Range($ws.Cells(2,$col),$ws.Cells(1+$unionSorted.Count,$col)).NumberFormatLocal='mm/dd hh:mm' }catch{}
    try{ $ws.Range($ws.Cells(2,$col+1),$ws.Cells(1+$unionSorted.Count,$col+1)).NumberFormatLocal='0.0' }catch{}

    # グラフ作成
    $ch=$ws.ChartObjects().Add(320,10,900,330); $c=$ch.Chart
    $c.ChartType=$xlXYScatterLines; $c.HasTitle=$true
    $c.ChartTitle.Text=("{0} - RTT by Floor" -f $t.Label)
    $c.Legend.Position=$xlLegendBottom
    try{ $c.SeriesCollection().Delete() }catch{}

    $colIdx=1
    foreach($fl in $floorList){
      $s=$c.SeriesCollection().NewSeries()
      $s.Name=("RTT ({0})" -f $fl)
      $endRowTime = ($ws.Cells($ws.Rows.Count,$colIdx).End(-4162)).Row  # xlUp=-4162
      $endRowVal  = ($ws.Cells($ws.Rows.Count,$colIdx+1).End(-4162)).Row
      $endRow = [Math]::Max($endRowTime,$endRowVal)
      $s.XValues=$ws.Range(($ws.Cells(2,$colIdx).Address()+":"+$ws.Cells($endRow,$colIdx).Address()))
      $s.Values =$ws.Range(($ws.Cells(2,$colIdx+1).Address()+":"+$ws.Cells($endRow,$colIdx+1).Address()))
      $colIdx+=2
    }
    # 閾値
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

    $created += $ws.Name

    # 大きなシートを作るたびに軽くGC（メモリ安定化）
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  # ---- INDEX ----
  $wsIdx=$wb.Worksheets.Add(); $wsIdx.Name='INDEX'
  $wsIdx.Cells(1,1).Value2='Sheets'
  $r=2
  foreach($w in @('LayerSeries','DeltaSeries')){
    $exists=$false
    for($ix=1;$ix -le $wb.Worksheets.Count;$ix++){ if($wb.Worksheets.Item($ix).Name -eq $w){ $exists=$true; break } }
    if($exists){ $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$w'!A1",'',$w) | Out-Null; $r++ }
  }
  foreach($sn in $created){ $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$sn'!A1",'',$sn) | Out-Null; $r++ }

  # 保存
  $wb.SaveAs($Output)
  Write-Host "Output: $Output"
}
catch{
  Write-Error (Format-Err $_)
  throw
}
finally{
  try{ Set-ExcelPerfMode $excel $false }catch{}
  if($wb){ try{ $wb.Close($false) }catch{}; Release-Com $wb; $wb=$null }
  if($excel){ try{ $excel.Quit() }catch{}; Release-Com $excel; $excel=$null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}