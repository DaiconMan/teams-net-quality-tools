<#
  Generate-NetQuality-HTMLReport.ps1 (PS 5.1 Compatible)
  - 入力CSV（文字コード）
    * teams_net_quality.csv（UTF-8 BOM）
      timestamp,target,hop_index,hop_ip,icmp_avg_ms,icmp_jitter_ms,loss_pct,notes,conn_type,ssid,bssid,signal_pct,ap_name,roamed,roam_from,roam_to,host,dns_ms,tcp_443_ms,http_head_ms,mos_estimate,probe,machine,user,tz_offset,source_file
    * targets.csv（UTF-8） … ヘッダー: role,key,label  (key=FQDN or IP)
      role: L2 / L3 / SAAS / RTR_WAN / RTR_LAN
    * node_roles.csv（UTF-8） … ヘッダー: ip_of_host,role,label,segment（任意）
    * floors.csv（UTF-8） … ヘッダー: bssid,area,floor,tag
  - 処理ポリシー
    * **宛先の決定は hop_ip を最優先**（targets.csv の key と一致したら採用）
      → 一致しなければ target を判定。**どちらにも一致しなければ除外**
    * 役割/ラベルは targets.csv を最優先。なければ node_roles.csv を補完
    * floors.csv の bssid を 12桁hex に正規化して area/floor/tag を付与（完全一致）
  - 出力: HTML単一ファイル（UTF-8 BOM, 外部ライブラリ不要）
  - 注意: 三項演算子( ?: )不使用 / $Host未使用 / OneDrive・日本語パス対応
#>

[CmdletBinding()]
param(
  [string]$QualityCsv = ".\teams_net_quality.csv",
  [string]$TargetsCsv = ".\targets.csv",
  [string]$NodeRoleCsv = ".\node_roles.csv",
  [string]$FloorsCsv = ".\floors.csv",
  [string]$OutHtml = ".\NetQuality-Report.html"
)

# ===== ヘルパ =====
function Parse-Double {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = ($s -replace '[^0-9\.\-]', '')
  $val = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)) {
    return [double]$val
  }
  return $null
}

function Get-Median { param([double[]]$arr)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  $a = [double]$s[$n/2 - 1]; $b = [double]$s[$n/2]
  return ($a + $b) / 2.0
}

function Get-Percentile { param([double[]]$arr, [double]$p)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  $k = ($n - 1) * $p
  $f = [math]::Floor($k); $c = [math]::Ceiling($k)
  if ($f -eq $c) { return [double]$s[[int]$k] }
  [double]$sf = $s[[int]$f]; [double]$sc = $s[[int]$c]
  return $sf + ($k - $f) * ($sc - $sf)
}

function Safe-Lower { param([string]$s)
  if ($null -eq $s) { return "" }
  return ($s.ToString().Trim().ToLower())
}

function Normalize-Bssid { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $h = ($s -replace '[^0-9a-fA-F]', '').ToLower()
  if ($h.Length -ge 12) { return $h.Substring(0,12) }
  return $h
}

function Import-CsvUtf8 { param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @() }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8
  if ($lines -is [string]) { $lines = @($lines) }
  if ($lines.Count -eq 0) { return @() }
  return $lines | ConvertFrom-Csv
}

# ===== CSV 読み込み =====
if (-not (Test-Path -LiteralPath $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $TargetsCsv)) { Write-Error "TargetsCsv が見つかりません: $TargetsCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $FloorsCsv))  { Write-Error "FloorsCsv が見つかりません: $FloorsCsv"; exit 1 }

$teams   = Import-CsvUtf8 -Path $QualityCsv
$targets = Import-CsvUtf8 -Path $TargetsCsv
$roles   = @()
if (Test-Path -LiteralPath $NodeRoleCsv) { $roles = Import-CsvUtf8 -Path $NodeRoleCsv }
$floors  = Import-CsvUtf8 -Path $FloorsCsv

# ===== targets.csv を辞書化 =====
$targetSet     = @{}   # key(lower) => $true
$roleByKey     = @{}   # key(lower) => role
$labelByKey    = @{}   # key(lower) => label
foreach($t in $targets){
  $k = Safe-Lower $t.key
  if ([string]::IsNullOrWhiteSpace($k)) { continue }
  $targetSet[$k] = $true
  if (-not [string]::IsNullOrWhiteSpace($t.role))  { $roleByKey[$k]  = $t.role }
  if (-not [string]::IsNullOrWhiteSpace($t.label)) { $labelByKey[$k] = $t.label }
}

# node_roles.csv（補助）
$roleByNode   = @{}
$labelByNode  = @{}
$segmentByNode= @{}
foreach($r in $roles){
  $nk = Safe-Lower $r.ip_of_host
  if ([string]::IsNullOrWhiteSpace($nk)) { continue }
  if (-not [string]::IsNullOrWhiteSpace($r.role))    { $roleByNode[$nk]   = $r.role }
  if (-not [string]::IsNullOrWhiteSpace($r.label))   { $labelByNode[$nk]  = $r.label }
  if (-not [string]::IsNullOrWhiteSpace($r.segment)) { $segmentByNode[$nk]= $r.segment }
}

# floors.csv（BSSID完全一致）
$areaByBssid  = @{}
$floorByBssid = @{}
$tagByBssid   = @{}
foreach($f in $floors){
  $b = Normalize-Bssid $f.bssid
  if ([string]::IsNullOrWhiteSpace($b)) { continue }
  if (-not $areaByBssid.ContainsKey($b))  { $areaByBssid[$b]  = $f.area }
  if (-not [string]::IsNullOrWhiteSpace($f.floor)) { $floorByBssid[$b] = $f.floor }
  if (-not [string]::IsNullOrWhiteSpace($f.tag))   { $tagByBssid[$b]   = $f.tag }
}

# ===== 正規化 & マッチング（hop_ip を優先）=====
$qual = @()
$cntMatchHop = 0; $cntMatchTarget = 0; $cntDropped = 0
$emptyBssid = 0; $emptyAp = 0

foreach($q in $teams){
  $hopKey = Safe-Lower $q.hop_ip
  $tgtKey = Safe-Lower $q.target

  $matchKey = $null
  $matchRole = $null
  $matchLabel = $null
  $segmentVal = ""

  # hop_ip → target の順で一致判定
  if (-not [string]::IsNullOrWhiteSpace($hopKey) -and $targetSet.ContainsKey($hopKey)) {
    $matchKey = $hopKey
    if ($roleByKey.ContainsKey($hopKey))  { $matchRole  = $roleByKey[$hopKey] }
    if ($labelByKey.ContainsKey($hopKey)) { $matchLabel = $labelByKey[$hopKey] }
    if ($segmentByNode.ContainsKey($hopKey)) { $segmentVal = $segmentByNode[$hopKey] }
    $cntMatchHop++
  } elseif (-not [string]::IsNullOrWhiteSpace($tgtKey) -and $targetSet.ContainsKey($tgtKey)) {
    $matchKey = $tgtKey
    if ($roleByKey.ContainsKey($tgtKey))  { $matchRole  = $roleByKey[$tgtKey] }
    if ($labelByKey.ContainsKey($tgtKey)) { $matchLabel = $labelByKey[$tgtKey] }
    if ($segmentByNode.ContainsKey($tgtKey)) { $segmentVal = $segmentByNode[$tgtKey] }
    $cntMatchTarget++
  } else {
    $cntDropped++
    continue
  }

  # 役割/ラベルの最終補完（targetsに無い時だけ node_roles を見る）
  if ([string]::IsNullOrWhiteSpace($matchRole) -and $roleByNode.ContainsKey($matchKey))  { $matchRole  = $roleByNode[$matchKey] }
  if ([string]::IsNullOrWhiteSpace($matchLabel) -and $labelByNode.ContainsKey($matchKey)){ $matchLabel = $labelByNode[$matchKey] }

  # floors: BSSID → area/floor/tag
  $bNorm = Normalize-Bssid $q.bssid
  if ($null -eq $bNorm) { $emptyBssid++ }
  $apName = $q.ap_name
  if ([string]::IsNullOrWhiteSpace($apName)) { $emptyAp++ }

  $areaVal = "Unknown"; $floorVal = $null; $tagVal = $null
  if ($bNorm -and $areaByBssid.ContainsKey($bNorm)) { $areaVal = $areaByBssid[$bNorm] }
  if ($bNorm -and $floorByBssid.ContainsKey($bNorm)) { $floorVal = $floorByBssid[$bNorm] }
  if ($bNorm -and $tagByBssid.ContainsKey($bNorm))   { $tagVal   = $tagByBssid[$bNorm] }

  # 数値系
  $rtt  = Parse-Double $q.icmp_avg_ms
  $jit  = Parse-Double $q.icmp_jitter_ms
  $loss = Parse-Double $q.loss_pct
  $mos  = Parse-Double $q.mos_estimate
  if ($null -eq $mos -and $null -ne $rtt -and $null -ne $loss) {
    $mos = [math]::Round((4.5 - 0.0004*[double]$rtt - 0.1*[double]$loss),2)
  }

  $obj = [PSCustomObject]@{
    timestamp = $q.timestamp
    target    = $matchKey      # ← 実際にマッチしたキー（hop_ip または target）
    rtt_ms    = $rtt
    jitter_ms = $jit
    loss_pct  = $loss
    mos       = $mos
    ssid      = $q.ssid
    bssid     = $bNorm
    ap_name   = $apName
    area      = $areaVal
    floor     = $floorVal
    ap_tag    = $tagVal
    role      = if ($matchRole) { $matchRole } else { "Uncategorized" }
    label     = if ($matchLabel){ $matchLabel } else { $matchKey }
    segment   = $segmentVal
  }
  $qual += $obj
}

# ===== 集計（エリア / AP / 対象 / 役割）=====
$summaryRows = @()
$groups = $qual | Group-Object -Property area, ap_name, target, role, segment
foreach($g in $groups){
  $rtts = @($g.Group | Where-Object {$_.rtt_ms -ne $null}    | ForEach-Object {[double]$_.rtt_ms})
  $jits = @($g.Group | Where-Object {$_.jitter_ms -ne $null} | ForEach-Object {[double]$_.jitter_ms})
  $loss = @($g.Group | Where-Object {$_.loss_pct -ne $null}  | ForEach-Object {[double]$_.loss_pct})
  $mosv = @($g.Group | Where-Object {$_.mos -ne $null}       | ForEach-Object {[double]$_.mos})

  $rtt_med = $null; if($rtts.Count -gt 0){ $rtt_med = [math]::Round((Get-Median $rtts),1) }
  $rtt_p95 = $null; if($rtts.Count -gt 0){ $rtt_p95 = [math]::Round((Get-Percentile $rtts 0.95),1) }
  $jit_med = $null; if($jits.Count -gt 0){ $jit_med = [math]::Round((Get-Median $jits),1) }
  $loss_avg= $null; if($loss.Count -gt 0){ $loss_avg = [math]::Round(($loss | Measure-Object -Average | Select-Object -ExpandProperty Average),2) }
  $mos_med = $null; if($mosv.Count -gt 0){ $mos_med = [math]::Round((Get-Median $mosv),2) }

  $summaryRows += [PSCustomObject]@{
    area     = $g.Group[0].area
    ap_name  = $g.Group[0].ap_name
    target   = $g.Group[0].target
    role     = $g.Group[0].role
    segment  = $g.Group[0].segment
    count    = $g.Count
    rtt_med  = $rtt_med
    rtt_p95  = $rtt_p95
    jit_med  = $jit_med
    loss_avg = $loss_avg
    mos_med  = $mos_med
  }
}

# ===== HTML（@'…'@ + JSON置換, UTF-8 BOM出力）=====
$summaryJson = $summaryRows | ConvertTo-Json -Depth 5

$htmlTemplate = @'
<!doctype html>
<html lang="ja"><head>
<meta charset="utf-8" />
<title>NetQuality Report (Targets+Hop Match)</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif; margin: 16px; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  .filters { display:flex; gap:8px; flex-wrap: wrap; margin: 8px 0 16px; }
  select, input, label { padding:6px; border:1px solid #ccc; border-radius: 8px; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 24px; }
  th, td { border-bottom: 1px solid #eee; padding: 8px; text-align: left; }
  th { background: #fafafa; position: sticky; top:0; z-index: 1; }
  .mono { font-family: ui-monospace, Menlo, Consolas, "Liberation Mono", monospace; }
  .rt-ok{background:#e7f7e7;} .rt-warn{background:#fff5e0;} .rt-bad{background:#fdecec;} .loss-bad{background:#fdecec;}
</style>
</head><body>
  <h1>ネットワーク品質レポート（targets.csv × hop_ip優先照合）</h1>

  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="targetSearch" placeholder="対象(部分一致)" />
    <select id="roleSel"><option value="">(すべての役割)</option></select>
    <select id="segSel"><option value="">(すべてのセグメント)</option></select>
  </div>

  <table id="sumTbl">
    <thead><tr>
      <th>エリア</th><th>AP</th><th>対象(キー)</th><th>役割</th><th>セグメント</th>
      <th>試行数</th><th>RTT中央値</th><th>RTT P95</th><th>ジッタ中央値</th><th>損失率(平均)</th><th>MOS(中央値)</th>
    </tr></thead><tbody></tbody>
  </table>

<script>
var summaryRows = __SUMMARY_JSON__;

function uniq(a){var o=[],i;for(i=0;i<a.length;i++){var x=a[i];if(x&&o.indexOf(x)===-1)o.push(x);}o.sort();return o;}
function fillSel(el,opts){for(var i=0;i<opts.length;i++){var op=document.createElement('option');op.textContent=opts[i];op.value=opts[i];el.appendChild(op);}}
var areaSel=document.getElementById('areaSel'),apSel=document.getElementById('apSel'),roleSel=document.getElementById('roleSel'),segSel=document.getElementById('segSel'),qInput=document.getElementById('targetSearch');
fillSel(areaSel,uniq(summaryRows.map(function(r){return r.area;})));
fillSel(apSel,uniq(summaryRows.map(function(r){return r.ap_name;}).filter(function(x){return !!x;})));
fillSel(roleSel,uniq(summaryRows.map(function(r){return r.role;})));
fillSel(segSel,uniq(summaryRows.map(function(r){return r.segment;})));
[areaSel,apSel,roleSel,segSel,qInput].forEach(function(el){el.addEventListener('input',render);el.addEventListener('change',render);});

function colorRtt(td,v){if(v==null)return;if(v<50)td.classList.add('rt-ok');else if(v<100)td.classList.add('rt-warn');else td.classList.add('rt-bad');}
function colorLoss(td,v){if(v!=null&&v>3)td.classList.add('loss-bad');}

function render(){
  var area=areaSel.value||"",ap=apSel.value||"",role=roleSel.value||"",seg=segSel.value||"",q=(qInput.value||"").toLowerCase();
  var tbody=document.querySelector('#sumTbl tbody'); tbody.innerHTML='';
  var rows=summaryRows.slice().sort(function(a,b){
    var ka=(a.area||"")+"|"+(a.ap_name||""); var kb=(b.area||"")+"|"+(b.ap_name||"");
    if(ka<kb)return-1;if(ka>kb)return 1; var ra=-1;if(a.rtt_med!=null)ra=-a.rtt_med; var rb=-1;if(b.rtt_med!=null)rb=-b.rtt_med; return ra-rb;
  });
  for(var i=0;i<rows.length;i++){
    var r=rows[i];
    if(area && r.area!==area)continue;
    if(ap && (r.ap_name||"")!==ap)continue;
    if(role && r.role!==role)continue;
    if(seg && r.segment!==seg)continue;
    if(q && String(r.target||"").toLowerCase().indexOf(q)===-1)continue;

    var tr=document.createElement('tr');
    function td(t){var e=document.createElement('td'); e.textContent=(t==null?"":t); return e;}

    tr.appendChild(td(r.area));
    tr.appendChild(td(r.ap_name||""));

    var ttd=td(r.target); ttd.classList.add('mono'); tr.appendChild(ttd);

    tr.appendChild(td(r.role||""));
    tr.appendChild(td(r.segment||""));
    tr.appendChild(td(r.count));

    var rttm=td(r.rtt_med); colorRtt(rttm,r.rtt_med); tr.appendChild(rttm);
    var rttp=td(r.rtt_p95); colorRtt(rttp,r.rtt_p95); tr.appendChild(rttp);
    tr.appendChild(td(r.jit_med));
    var loss=td(r.loss_avg); colorLoss(loss,r.loss_avg); tr.appendChild(loss);
    tr.appendChild(td(r.mos_med));

    tbody.appendChild(tr);
  }
}
render();
</script>
</body></html>
'@

$html = $htmlTemplate.Replace('__SUMMARY_JSON__', $summaryJson)

# ===== 出力 (UTF-8 BOM) =====
try {
  $fullPath = [System.IO.Path]::GetFullPath($OutHtml)
  $dir = [System.IO.Path]::GetDirectoryName($fullPath)
  if (-not [string]::IsNullOrWhiteSpace($dir)) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
  }
  $enc = New-Object System.Text.UTF8Encoding($true) # BOM付き
  [System.IO.File]::WriteAllText($fullPath, $html, $enc)
  Write-Output ("HTMLレポートを出力しました(UTF-8 BOM): {0}" -f $fullPath)
} catch {
  $html | Out-File -FilePath $OutHtml
  Write-Warning "WriteAllText に失敗したため Out-File で出力しました。"
}

# ===== 参考メトリクス（数値のみ表示）=====
$total = $teams.Count
$after = $qual.Count
$mappedArea = ($qual | Where-Object { $_.area -ne "Unknown" }).Count
if ($total -gt 0) {
  $pctUsed = [math]::Round(100.0 * $after / $total, 1)
  $pctArea = 0.0
  if ($after -gt 0) { $pctArea = [math]::Round(100.0 * $mappedArea / $after, 1) }
  Write-Output ("teams行数: {0}, 採用行数( targets×(hop_ip→target) ): {1} ({2}%)" -f $total,$after,$pctUsed)
  Write-Output ("area付与率（Unknown除外）: {0}% ({1}/{2})" -f $pctArea,$mappedArea,$after)
  Write-Output ("hop_ipで一致: {0} / targetで一致: {1} / 除外: {2}" -f $cntMatchHop,$cntMatchTarget,$cntDropped)
  Write-Output ("bssid空欄: {0} / ap_name空欄: {1}" -f $emptyBssid,$emptyAp)
}
