<#
  Generate-NetQuality-HTMLReport.ps1 (PS 5.1 Compatible)
  - 入力: teams_net_quality.csv
  - 補正: floor.csv (BSSID/SSID → area/floor/tag), node_roles.csv (IP/FQDN → role/label/segment)
  - 出力: HTML 単一ファイル（UTF-8 BOM）
  - 仕様: path_hop_quality.csv は不使用（Zscaler配慮）
  - 既定で SaaS/Internet を除外（UIトグルで含め可能）
  - 列名：BOM/空白/大小文字/別名を自動吸収
  - OneDrive/日本語パス対応（出力先ディレクトリ自動作成）
  - 注意: PowerShellの $Host は未使用。CSVの列名 host は target として扱う。
#>

[CmdletBinding()]
param(
  [string]$QualityCsv    = ".\teams_net_quality.csv",
  [string]$BssidFloorCsv = ".\floor.csv",
  [string]$NodeRoleCsv   = ".\node_roles.csv",
  [string]$OutHtml       = ".\NetQuality-Report.html"
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

function Get-Median {
  param([double[]]$arr)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  else {
    $a = [double]$s[$n/2 - 1]; $b = [double]$s[$n/2]
    return ($a + $b) / 2.0
  }
}

function Get-Percentile {
  param([double[]]$arr, [double]$p) # 0.95 等
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  $k = ($n - 1) * $p
  $f = [math]::Floor($k); $c = [math]::Ceiling($k)
  if ($f -eq $c) { return [double]$s[[int]$k] }
  [double]$sf = $s[[int]$f]; [double]$sc = $s[[int]$c]
  return $sf + ($k - $f) * ($sc - $sf)
}

function Safe-Lower {
  param([string]$s)
  if ($null -eq $s) { return "" }
  return ($s.ToString().Trim().ToLower())
}

function Normalize-Bssid {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $h = ($s -replace '[^0-9a-fA-F]', '').ToLower()
  if ($h.Length -ge 12) { return $h.Substring(0,12) }
  return $h
}

function Normalize-Ssid {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  return ($s.Trim().ToLower())
}

# 列名のBOM/空白/大小文字揺れを吸収
function Normalize-HeaderName {
  param([string]$name)
  if ($null -eq $name) { return "" }
  $n = $name.Replace([string]([char]0xFEFF), '') # BOM除去
  $n = $n.Trim().ToLower()
  return $n
}

# Import-Csv の安全版（UTF8優先、失敗時デフォルト）
function Import-CsvSafe {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @() }
  try {
    return Import-Csv -Path $Path -Encoding UTF8
  } catch {
    return Import-Csv -Path $Path
  }
}

# floor の bssid/ssid が接頭辞（末尾 *）かを判定して正規化
function Get-MapKey {
  param([string]$raw, [switch]$IsSsid)
  if ([string]::IsNullOrWhiteSpace($raw)) {
    return [PSCustomObject]@{ Key=$null; IsPrefix=$false }
  }
  $trimmed = $raw.Trim()
  $isPrefix = $false
  if ($trimmed.EndsWith("*")) {
    $isPrefix = $true
    $trimmed = $trimmed.Substring(0, $trimmed.Length - 1)
  }
  if ($IsSsid) { $norm = Normalize-Ssid $trimmed } else { $norm = Normalize-Bssid $trimmed }
  return [PSCustomObject]@{ Key=$norm; IsPrefix=$isPrefix }
}

# ===== CSV 読み込み =====
if (-not (Test-Path -LiteralPath $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
$qualRaw  = Import-CsvSafe -Path $QualityCsv
$floorMap = @()
if (Test-Path -LiteralPath $BssidFloorCsv) { $floorMap = Import-CsvSafe -Path $BssidFloorCsv }
$nodeRoles = @()
if (Test-Path -LiteralPath $NodeRoleCsv)   { $nodeRoles = Import-CsvSafe -Path $NodeRoleCsv }

# ===== 列名の自動検出（強化）=====
function Find-Col {
  param(
    [object]$row,
    [string[]]$candidates
  )
  $rawProps = @($row.PSObject.Properties.Name)
  $props = @()
  foreach($p in $rawProps){
    $props += (Normalize-HeaderName $p)
  }

  # 1) 完全一致
  foreach($cand in $candidates){
    $c = Normalize-HeaderName $cand
    for($i=0; $i -lt $props.Count; $i++){
      if ($props[$i] -eq $c) { return $rawProps[$i] }
    }
  }

  # 2) 単語境界一致（^cand$ | ^cand[_\-] | [_\-]cand[_\-] | [_\-]cand$）
  foreach($cand in $candidates){
    $c = Normalize-HeaderName $cand
    for($i=0; $i -lt $props.Count; $i++){
      $pn = $props[$i]
      $hit = $false
      if ($pn -eq $c) { $hit = $true }
      else {
        if ($pn.StartsWith($c + "_") -or $pn.StartsWith($c + "-")) { $hit = $true }
        elseif ($pn.EndsWith("_" + $c) -or $pn.EndsWith("-" + $c)) { $hit = $true }
        elseif ($pn.Contains("_" + $c + "_") -or $pn.Contains("-" + $c + "-")) { $hit = $true }
      }
      if ($hit) { return $rawProps[$i] }
    }
  }

  # 3) 部分一致（最後の手段）
  foreach($cand in $candidates){
    $c = Normalize-HeaderName $cand
    for($i=0; $i -lt $props.Count; $i++){
      if ($props[$i].Contains($c)) { return $rawProps[$i] }
    }
  }
  return $null
}

function Get-Val {
  param($row, $colName)
  if ([string]::IsNullOrWhiteSpace($colName)) { return $null }
  $val = $row.$colName
  if ($null -eq $val) { return $null }
  $txt = $val.ToString()
  # 先頭BOM等を除去
  $txt = $txt.Replace([string]([char]0xFEFF), '')
  $txt = $txt.Trim()
  if ($txt -eq "") { return $null }
  return $txt
}

# 先頭行から列を推定
$probe = $null
if ($qualRaw.Count -gt 0) { $probe = $qualRaw[0] } else { Write-Error "QualityCsv にデータがありません"; exit 1 }

$colMap = @{
  timestamp = Find-Col $probe @('timestamp','time','datetime','date','collected_at')
  target    = Find-Col $probe @('host','target','dest','destination','fqdn','dnsname','dns','ip','address','addr')
  rtt_ms    = Find-Col $probe @('icmp_avg_ms','avg_ms','rtt_ms','avg_rtt_ms','ping_avg_ms','latency_ms','avg_latency_ms','avg_rtt','rtt')
  jitter_ms = Find-Col $probe @('icmp_jitter_ms','jitter_ms','jitter','avg_jitter_ms')
  loss_pct  = Find-Col $probe @('loss_pct','packet_loss_pct','loss_percent','loss_rate_pct','loss','packet_loss')
  mos       = Find-Col $probe @('mos','mean_opinion_score')
  ssid      = Find-Col $probe @('ssid','wifi_ssid','wlan_ssid')
  bssid     = Find-Col $probe @('bssid','wifi_bssid','wlan_bssid','connected_bssid')
  ap_name   = Find-Col $probe @('ap_name','ap','apname','access_point','connected_ap','ap_hostname')
}

Write-Output "列マッピング: $(($colMap.GetEnumerator() | Sort-Object Name | ForEach-Object { '{0}→{1}' -f $_.Name, ($_.Value -as [string]) }) -join ', ')"

# ===== floor.csv マップ作成（BSSID/SSID 両対応）=====
$areaByBssidExact  = @{}
$floorByBssidExact = @{}
$tagByBssidExact   = @{}
$prefixBssid = @()  # @{ Prefix="aabbcc"; Area="X"; Floor="Y"; Tag="Z" }

$areaBySsidExact  = @{}
$floorBySsidExact = @{}
$tagBySsidExact   = @{}
$prefixSsid = @()  # @{ Prefix="corp-wifi"; Area="X"; Floor="Y"; Tag="Z" }

foreach($r in $floorMap){
  # 列名BOM対策
  $areaCol  = 'area';  $floorCol = 'floor'; $tagCol = 'tag'
  $bssidCol = 'bssid'; $ssidCol  = 'ssid'
  $props = @($r.PSObject.Properties.Name)
  foreach($p in $props){
    $np = Normalize-HeaderName $p
    if ($np -eq 'area')  { $areaCol  = $p }
    if ($np -eq 'floor') { $floorCol = $p }
    if ($np -eq 'tag')   { $tagCol   = $p }
    if ($np -eq 'bssid') { $bssidCol = $p }
    if ($np -eq 'ssid')  { $ssidCol  = $p }
  }

  $area = Get-Val $r $areaCol
  $floor= Get-Val $r $floorCol
  $tag  = Get-Val $r $tagCol

  # BSSIDキー
  $bssidRaw = Get-Val $r $bssidCol
  if ($null -ne $bssidRaw) {
    $binfo = Get-MapKey -raw $bssidRaw
    if ($binfo.Key) {
      if ($binfo.IsPrefix -or ($binfo.Key.Length -lt 12)) {
        $prefixBssid += @{ Prefix=$binfo.Key; Area=$area; Floor=$floor; Tag=$tag }
      } else {
        $areaByBssidExact[$binfo.Key]  = $area
        if ($null -ne $floor) { $floorByBssidExact[$binfo.Key] = $floor }
        if ($null -ne $tag)   { $tagByBssidExact[$binfo.Key]   = $tag }
      }
    }
  }

  # SSIDキー
  $ssidRaw = Get-Val $r $ssidCol
  if ($null -ne $ssidRaw) {
    $sinfo = Get-MapKey -raw $ssidRaw -IsSsid
    if ($sinfo.Key) {
      if ($sinfo.IsPrefix) {
        $prefixSsid += @{ Prefix=$sinfo.Key; Area=$area; Floor=$floor; Tag=$tag }
      } else {
        $areaBySsidExact[$sinfo.Key]  = $area
        if ($null -ne $floor) { $floorBySsidExact[$sinfo.Key] = $floor }
        if ($null -ne $tag)   { $tagBySsidExact[$sinfo.Key]   = $tag }
      }
    }
  }
}

# マッチ内訳カウンタ
$cntMatchBssidExact = 0
$cntMatchBssidPref  = 0
$cntMatchSsidExact  = 0
$cntMatchSsidPref   = 0
$cntMatchNone       = 0

function Lookup-Meta {
  param([string]$bssidRaw, [string]$ssidRaw)
  $area="Unknown"; $floor=$null; $tag=$null; $bKey=$null

  $normB = Normalize-Bssid $bssidRaw
  $normS = Normalize-Ssid $ssidRaw

  # 1) BSSID 完全一致
  if ($normB -and $areaByBssidExact.ContainsKey($normB)) {
    $area = $areaByBssidExact[$normB]
    if ($floorByBssidExact.ContainsKey($normB)) { $floor = $floorByBssidExact[$normB] }
    if ($tagByBssidExact.ContainsKey($normB))   { $tag   = $tagByBssidExact[$normB] }
    $bKey = $normB; $script:cntMatchBssidExact++
    return [PSCustomObject]@{ Area=$area; Floor=$floor; Tag=$tag; Key=$bKey }
  }

  # 2) BSSID 接頭辞
  if ($normB -and $prefixBssid.Count -gt 0) {
    foreach($p in $prefixBssid){
      $pref = $p.Prefix
      if ($pref -and $normB.StartsWith($pref)) {
        $area = $p.Area; $floor=$p.Floor; $tag=$p.Tag; $bKey=$normB; $script:cntMatchBssidPref++
        return [PSCustomObject]@{ Area=$area; Floor=$floor; Tag=$tag; Key=$bKey }
      }
    }
  }

  # 3) SSID 完全一致
  if ($normS -and $areaBySsidExact.ContainsKey($normS)) {
    $area = $areaBySsidExact[$normS]
    if ($floorBySsidExact.ContainsKey($normS)) { $floor = $floorBySsidExact[$normS] }
    if ($tagBySsidExact.ContainsKey($normS))   { $tag   = $tagBySsidExact[$normS] }
    $script:cntMatchSsidExact++
    return [PSCustomObject]@{ Area=$area; Floor=$floor; Tag=$tag; Key=$normB }
  }

  # 4) SSID 接頭辞
  if ($normS -and $prefixSsid.Count -gt 0) {
    foreach($p in $prefixSsid){
      $pref = $p.Prefix
      if ($pref -and $normS.StartsWith($pref)) {
        $area = $p.Area; $floor=$p.Floor; $tag=$p.Tag; $script:cntMatchSsidPref++
        return [PSCustomObject]@{ Area=$area; Floor=$floor; Tag=$tag; Key=$normB }
      }
    }
  }

  $script:cntMatchNone++
  return [PSCustomObject]@{ Area=$area; Floor=$floor; Tag=$tag; Key=$normB }
}

# ノード（IP/FQDN）→役割/ラベル/セグメント
$roleByNode    = @{}
$labelByNode   = @{}
$segmentByNode = @{}
foreach($r in $nodeRoles){
  $props = @($r.PSObject.Properties.Name)
  $ipCol = $null; $roleCol=$null; $labelCol=$null; $segCol=$null
  foreach($p in $props){
    $np = Normalize-HeaderName $p
    if ($np -eq 'ip_or_host' -or $np -eq 'host' -or $np -eq 'ip') { $ipCol = $p }
    if ($np -eq 'role')   { $roleCol  = $p }
    if ($np -eq 'label')  { $labelCol = $p }
    if ($np -eq 'segment'){ $segCol   = $p }
  }
  $k = Safe-Lower (Get-Val $r $ipCol)
  if (-not [string]::IsNullOrWhiteSpace($k)) {
    $roleByNode[$k] = Get-Val $r $roleCol
    $lv = Get-Val $r $labelCol
    if ($lv) { $labelByNode[$k]   = $lv }
    $sv = Get-Val $r $segCol
    if ($sv) { $segmentByNode[$k] = $sv }
  }
}

# ===== 正規化（teams_net_quality）=====
$qual = @()
$emptyTarget = 0; $emptyBssid = 0; $emptyAp = 0

foreach($q in $qualRaw){
  $targetTxt = Get-Val $q $colMap.target
  if ($null -eq $targetTxt) { $emptyTarget++ }

  $bssidRaw = Get-Val $q $colMap.bssid
  if ($null -eq $bssidRaw) { $emptyBssid++ }

  $ssidRaw  = Get-Val $q $colMap.ssid
  $meta = Lookup-Meta $bssidRaw $ssidRaw

  $apName = Get-Val $q $colMap.ap_name
  if ($null -eq $apName) { $emptyAp++ }

  $rttTxt = Get-Val $q $colMap.rtt_ms
  $jitTxt = Get-Val $q $colMap.jitter_ms
  $losTxt = Get-Val $q $colMap.loss_pct
  $mosTxt = Get-Val $q $colMap.mos

  $obj = [PSCustomObject]@{
    timestamp = Get-Val $q $colMap.timestamp
    target    = $targetTxt
    rtt_ms    = Parse-Double $rttTxt
    jitter_ms = Parse-Double $jitTxt
    loss_pct  = Parse-Double $losTxt
    mos       = Parse-Double $mosTxt
    ssid      = $ssidRaw
    bssid     = $meta.Key
    ap_name   = $apName
    area      = $meta.Area
    floor     = $meta.Floor
    ap_tag    = $meta.Tag
  }

  if ($null -eq $obj.mos -and $null -ne $obj.rtt_ms -and $null -ne $obj.loss_pct) {
    $obj.mos = [math]::Round((4.5 - 0.0004*[double]$obj.rtt_ms - 0.1*[double]$obj.loss_pct),2)
  }

  $nk = Safe-Lower $obj.target
  $roleVal = "Uncategorized"
  if ($roleByNode.ContainsKey($nk)) { $roleVal = $roleByNode[$nk] }
  $obj | Add-Member -NotePropertyName role -NotePropertyValue $roleVal

  $labelVal = $obj.target
  if ($labelByNode.ContainsKey($nk)) { $labelVal = $labelByNode[$nk] }
  $obj | Add-Member -NotePropertyName label -NotePropertyValue $labelVal

  $segVal = ""
  if ($segmentByNode.ContainsKey($nk)) { $segVal = $segmentByNode[$nk] }
  $obj | Add-Member -NotePropertyName segment -NotePropertyValue $segVal

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

# ===== HTML テンプレ生成（@' … '@ + JSON置換）=====
$summaryJson = $summaryRows | ConvertTo-Json -Depth 5

$htmlTemplate = @'
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<title>NetQuality Report (No-Hop / Internal Focus)</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif; margin: 16px; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  .filters { display:flex; gap:8px; flex-wrap: wrap; margin: 8px 0 16px; }
  select, input, label { padding:6px; border:1px solid #ccc; border-radius: 8px; }
  label.chk { border:none; padding:0 6px 0 0; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 24px; }
  th, td { border-bottom: 1px solid #eee; padding: 8px; text-align: left; }
  th { background: #fafafa; position: sticky; top:0; z-index: 1; }
  .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
  .hint { color:#666; font-size:12px; }
  .legend { font-size:12px; color:#444; margin:4px 0 12px; }
  .legend span { margin-right:12px; padding:2px 6px; border-radius:6px; }
  .rt-ok   { background:#e7f7e7; }
  .rt-warn { background:#fff5e0; }
  .rt-bad  { background:#fdecec; }
  .loss-bad{ background:#fdecec; }
  .muted { color:#777; }
</style>
</head>
<body>
  <h1>ネットワーク品質レポート（Hop解析なし／内部重視）</h1>
  <div class="hint">
    このレポートは <strong>path_hop_quality.csv を使用しません</strong>。Zscaler経由でのInternet/SaaS向けpingは信頼しない想定のため、既定で表示から除外します。
  </div>

  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="targetSearch" placeholder="対象(部分一致)" />
    <select id="roleSel"><option value="">(すべての役割)</option></select>
    <select id="segSel"><option value="">(すべてのセグメント)</option></select>
    <label class="chk"><input type="checkbox" id="includeExternal"> 外部(SaaS/Internet)を含める</label>
  </div>

  <div class="legend">
    <span class="rt-ok">RTT &lt; 50ms</span>
    <span class="rt-warn">50–100ms</span>
    <span class="rt-bad">≥ 100ms</span>
    <span class="loss-bad">損失 &gt; 3%</span>
  </div>

  <table id="sumTbl">
    <thead>
      <tr>
        <th>エリア</th>
        <th>AP</th>
        <th>対象</th>
        <th>役割</th>
        <th>セグメント</th>
        <th>試行数</th>
        <th>RTT中央値</th>
        <th>RTT P95</th>
        <th>ジッタ中央値</th>
        <th>損失率(平均)</th>
        <th>MOS(中央値)</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>

<script>
  var summaryRows = __SUMMARY_JSON__;

  function uniq(vals){
    var out = [];
    for (var i=0;i<vals.length;i++){
      var x = vals[i];
      if (x && out.indexOf(x) === -1) { out.push(x); }
    }
    out.sort();
    return out;
  }
  function fillSel(el, opts){
    for (var i=0;i<opts.length;i++){
      var op = document.createElement('option');
      op.textContent = opts[i]; op.value = opts[i];
      el.appendChild(op);
    }
  }
  var areaSel = document.getElementById('areaSel');
  var apSel   = document.getElementById('apSel');
  var roleSel = document.getElementById('roleSel');
  var segSel  = document.getElementById('segSel');
  var qInput  = document.getElementById('targetSearch');
  var includeExternal = document.getElementById('includeExternal');

  fillSel(areaSel, uniq(summaryRows.map(function(r){return r.area;})));
  fillSel(apSel,   uniq(summaryRows.map(function(r){return r.ap_name;}).filter(function(x){return !!x;})));
  fillSel(roleSel, uniq(summaryRows.map(function(r){return r.role;})));
  fillSel(segSel,  uniq(summaryRows.map(function(r){return r.segment;})));

  function addInputHandler(el){ el.addEventListener('input', render); el.addEventListener('change', render); }
  addInputHandler(areaSel); addInputHandler(apSel); addInputHandler(roleSel);
  addInputHandler(segSel);  addInputHandler(qInput); addInputHandler(includeExternal);

  function colorRtt(td, val){
    if (val == null) return;
    if (val < 50) td.classList.add('rt-ok');
    else if (val < 100) td.classList.add('rt-warn');
    else td.classList.add('rt-bad');
  }
  function colorLoss(td, val){ if (val != null && val > 3) td.classList.add('loss-bad'); }
  function isExternalRole(role){
    if (!role) return false;
    var r = String(role).toLowerCase();
    if (r === 'saas') return true;
    if (r === 'internet') return true;
    return false;
  }

  function render(){
    var area = areaSel.value || "";
    var ap   = apSel.value || "";
    var role = roleSel.value || "";
    var seg  = segSel.value || "";
    var q    = (qInput.value || "").toLowerCase();
    var showExt = includeExternal.checked;

    var tbody = document.querySelector('#sumTbl tbody');
    tbody.innerHTML = '';

    var rows = summaryRows.slice().sort(function(a,b){
      var ka = (a.area||"") + "|" + (a.ap_name||"");
      var kb = (b.area||"") + "|" + (b.ap_name||"");
      if (ka < kb) return -1;
      if (ka > kb) return 1;
      var ra = -1; if (a.rtt_med != null) { ra = -a.rtt_med; }
      var rb = -1; if (b.rtt_med != null) { rb = -b.rtt_med; }
      return ra - rb;
    });

    for (var i=0;i<rows.length;i++){
      var r = rows[i];
      if (area && r.area !== area) continue;
      if (ap && (r.ap_name||"") !== ap) continue;
      if (role && r.role !== role) continue;
      if (seg && r.segment !== seg) continue;
      if (!showExt && isExternalRole(r.role)) continue;
      if (q && String(r.target||"").toLowerCase().indexOf(q) === -1) continue;

      var tr = document.createElement('tr');
      function td(t){ var e=document.createElement('td'); e.textContent = (t==null?"":t); return e; }

      tr.appendChild(td(r.area));
      tr.appendChild(td(r.ap_name||""));

      var ttd = td(r.target); ttd.classList.add('mono'); tr.appendChild(ttd);

      var rtd = document.createElement('td');
      rtd.textContent = (r.role||""); if (isExternalRole(r.role)) { rtd.classList.add('muted'); }
      tr.appendChild(rtd);

      tr.appendChild(td(r.segment||""));
      tr.appendChild(td(r.count));

      var rttm = td(r.rtt_med); colorRtt(rttm, r.rtt_med); tr.appendChild(rttm);
      var rttp = td(r.rtt_p95); colorRtt(rttp, r.rtt_p95); tr.appendChild(rttp);
      tr.appendChild(td(r.jit_med));
      var loss = td(r.loss_avg); colorLoss(loss, r.loss_avg); tr.appendChild(loss);
      tr.appendChild(td(r.mos_med));

      tbody.appendChild(tr);
    }
  }

  includeExternal.checked = false;
  render();
</script>
</body>
</html>
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
  $html | Out-File -FilePath $OutHtml -Encoding utf8
  Write-Warning "WriteAllText に失敗したため Out-File で出力しました。"
}

# ===== 参考メトリクス（数値のみ表示、値は出さない）=====
$total = $qual.Count
$mapped = ($qual | Where-Object { $_.area -ne "Unknown" }).Count
if ($total -gt 0) {
  $pct = [math]::Round(100.0 * $mapped / $total, 1)
  Write-Output ("BSSID/SSID→エリア マッピング率: {0}% ({1}/{2})" -f $pct,$mapped,$total)
  Write-Output ("target 空欄: {0} / bssid 空欄: {1} / ap_name 空欄: {2}" -f $emptyTarget,$emptyBssid,$emptyAp)
  Write-Output ("match内訳: BSSID完全={0}, BSSID接頭辞={1}, SSID完全={2}, SSID接頭辞={3}, 未マッチ={4}" -f `
    $cntMatchBssidExact, $cntMatchBssidPref, $cntMatchSsidExact, $cntMatchSsidPref, $cntMatchNone)
}
