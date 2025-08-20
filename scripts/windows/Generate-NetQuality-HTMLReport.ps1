<#
  Generate-NetQuality-HTMLReport.ps1
  - 既存の計測CSV(teams_net_quality.csv / path_hop_quality.csv)と補正CSV(floor.csv, node_roles.csv)を読み込み
  - エリア/ AP / 対象別に RTT中央値 / 95パーセンタイル / 損失率 / MOSを集計
  - Hopごとの差分からボトルネック候補を推定
  - 単一のHTMLレポートを出力（外部ライブラリ不要、PS5.1対応）

  使い方例:
    powershell -ExecutionPolicy Bypass -File .\Generate-NetQuality-HTMLReport.ps1 `
      -QualityCsv "$env:LOCALAPPDATA\TeamsNet\teams_net_quality.csv" `
      -HopCsv "$env:LOCALAPPDATA\TeamsNet\path_hop_quality.csv" `
      -BssidFloorCsv ".\floor.csv" -NodeRoleCsv ".\node_roles.csv" `
      -OutHtml ".\NetQuality-Report.html"
#>

[CmdletBinding()]
param(
  [string]$QualityCsv    = (Join-Path $env:LOCALAPPDATA "TeamsNet\teams_net_quality.csv"),
  [string]$HopCsv        = (Join-Path $env:LOCALAPPDATA "TeamsNet\path_hop_quality.csv"),
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
  param([double[]]$arr, [double]$p) # p: 0.95 等
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

# ===== CSV 読み込み =====
if (-not (Test-Path $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
$qualRaw = Import-Csv -Path $QualityCsv

$hopRaw = @()
if (Test-Path $HopCsv) { $hopRaw = Import-Csv -Path $HopCsv }

$floorMap = @()
if (Test-Path $BssidFloorCsv) { $floorMap = Import-Csv -Path $BssidFloorCsv }

$nodeRoles = @()
if (Test-Path $NodeRoleCsv) { $nodeRoles = Import-Csv -Path $NodeRoleCsv }

# BSSID→Area辞書
$areaByBssid = @{}
foreach($r in $floorMap){
  $b = Safe-Lower $r.bssid
  if (-not $areaByBssid.ContainsKey($b)) { $areaByBssid[$b] = $r.area }
}

# ホスト/アドレス→役割辞書
$roleByNode = @{}
$labelByNode = @{}
$segmentByNode = @{}
foreach($r in $nodeRoles){
  $k = Safe-Lower $r.ip_or_host
  $roleByNode[$k] = $r.role
  if ($r.PSObject.Properties.Name -contains 'label')   { $labelByNode[$k] = $r.label }
  if ($r.PSObject.Properties.Name -contains 'segment') { $segmentByNode[$k] = $r.segment }
}

# ===== 正規化（teams_net_quality）=====
$qual = @()
foreach($q in $qualRaw){
  $obj = [PSCustomObject]@{
    timestamp = $q.timestamp
    host      = $q.host
    rtt_ms    = Parse-Double $q.icmp_avg_ms
    jitter_ms = Parse-Double $q.icmp_jitter_ms
    loss_pct  = Parse-Double $q.loss_pct
    mos       = Parse-Double $q.mos
    ssid      = $q.ssid
    bssid     = Safe-Lower $q.bssid
    ap_name   = $q.ap_name
    area      = $null
  }
  if ($null -eq $obj.mos) { $obj.mos = 4.5 - 0.0004*([double]($obj.rtt_ms | ForEach-Object {$_})) - 0.1*([double]($obj.loss_pct | ForEach-Object {$_})) }
  if ($areaByBssid.ContainsKey($obj.bssid)) { $obj.area = $areaByBssid[$obj.bssid] } else { $obj.area = "Unknown" }
  $qual += $obj
}

# ホスト役割付与
foreach($q in $qual){
  $k = Safe-Lower $q.host
  $q | Add-Member -NotePropertyName role -NotePropertyValue ($roleByNode.ContainsKey($k) ? $roleByNode[$k] : "Uncategorized")
  $q | Add-Member -NotePropertyName node_label -NotePropertyValue ($labelByNode.ContainsKey($k) ? $labelByNode[$k] : $q.host)
  $q | Add-Member -NotePropertyName segment -NotePropertyValue ($segmentByNode.ContainsKey($k) ? $segmentByNode[$k] : "")
}

# ===== 集計（エリア/ AP / 対象）=====
$summaryRows = @()
$groups = $qual | Group-Object -Property area, ap_name, host, role
foreach($g in $groups){
  $rtts = @($g.Group | Where-Object {$_.rtt_ms -ne $null} | ForEach-Object {[double]$_.rtt_ms})
  $loss = @($g.Group | Where-Object {$_.loss_pct -ne $null} | ForEach-Object {[double]$_.loss_pct})
  $mosv = @($g.Group | Where-Object {$_.mos -ne $null} | ForEach-Object {[double]$_.mos})

  $item = [PSCustomObject]@{
    area     = $g.Group[0].area
    ap_name  = $g.Group[0].ap_name
    host     = $g.Group[0].host
    role     = $g.Group[0].role
    count    = $g.Count
    rtt_med  = if($rtts.Count -gt 0){ [math]::Round((Get-Median $rtts),1) } else { $null }
    rtt_p95  = if($rtts.Count -gt 0){ [math]::Round((Get-Percentile $rtts 0.95),1) } else { $null }
    loss_avg = if($loss.Count -gt 0){ [math]::Round(($loss | Measure-Object -Average | Select-Object -ExpandProperty Average),2) } else { $null }
    mos_med  = if($mosv.Count -gt 0){ [math]::Round((Get-Median $mosv),2) } else { $null }
  }
  $summaryRows += $item
}

# ===== Hop差分分析（ボトルネック推定）=====
$bottleneckRows = @()
if ($hopRaw.Count -gt 0) {
  # 正規化 & 不要行除外
  $hopNorm = @()
  foreach($h in $hopRaw){
    # hop_index が数値でない行（tracert_no_reply等）は除外
    $idx = $null
    if ([int]::TryParse(($h.hop_index), [ref]$idx)) {
      $obj = [PSCustomObject]@{
        timestamp = $h.timestamp
        target    = $h.target
        hop_index = $idx
        hop_ip    = $h.hop_ip
        rtt_ms    = Parse-Double $h.icmp_avg_ms
        bssid     = Safe-Lower $h.bssid
        ap_name   = $h.ap_name
      }
      $obj | Add-Member -NotePropertyName area -NotePropertyValue (($areaByBssid.ContainsKey($obj.bssid)) ? $areaByBssid[$obj.bssid] : "Unknown")
      $hopNorm += $obj
    }
  }

  # (timestamp,target,area) ごとに並べ替えて最大差分を抽出
  $paths = $hopNorm | Group-Object -Property timestamp, target, area
  $bncRaw = @()
  foreach($p in $paths){
    $rows = $p.Group | Sort-Object hop_index
    if ($rows.Count -lt 2) { continue }
    $maxDelta = $null; $ipAtMax = ""; $idxAtMax = $null; $hopRtt = $null
    for ($i=1; $i -lt $rows.Count; $i++){
      $prev = $rows[$i-1]; $cur = $rows[$i]
      if ($null -ne $prev.rtt_ms -and $null -ne $cur.rtt_ms) {
        $delta = [double]$cur.rtt_ms - [double]$prev.rtt_ms
        if ($delta -gt 0 -and ($null -eq $maxDelta -or $delta -gt $maxDelta)) {
          $maxDelta = $delta; $ipAtMax = $cur.hop_ip; $idxAtMax = $cur.hop_index; $hopRtt = $cur.rtt_ms
        }
      }
    }
    if ($null -ne $maxDelta) {
      $bnc = [PSCustomObject]@{
        area       = $rows[0].area
        target     = $rows[0].target
        ap_name    = $rows[0].ap_name
        hop_ip     = $ipAtMax
        hop_index  = $idxAtMax
        delta_ms   = [math]::Round($maxDelta,1)
        hop_rtt_ms = [math]::Round($hopRtt,1)
      }
      # 役割付与
      $k = Safe-Lower $bnc.hop_ip
      $bnc | Add-Member -NotePropertyName role -NotePropertyValue ($roleByNode.ContainsKey($k) ? $roleByNode[$k] : "Unknown")
      $bnc | Add-Member -NotePropertyName node_label -NotePropertyValue ($labelByNode.ContainsKey($k) ? $labelByNode[$k] : $bnc.hop_ip)
      $bncRaw += $bnc
    }
  }

  # (area,target)単位で「代表ボトルネック」を選定（delta_ms の中央値が最大の hop）
  $grp2 = $bncRaw | Group-Object -Property area, target, hop_ip, role, node_label
  $agg2 = @()
  foreach($g in $grp2){
    $deltas = @($g.Group | ForEach-Object {[double]$_.delta_ms})
    $agg2 += [PSCustomObject]@{
      area       = $g.Group[0].area
      target     = $g.Group[0].target
      hop_ip     = $g.Group[0].hop_ip
      role       = $g.Group[0].role
      node_label = $g.Group[0].node_label
      delta_med  = [math]::Round((Get-Median $deltas),1)
    }
  }
  $pick = @()
  $byAreaTarget = $agg2 | Group-Object -Property area, target
  foreach($g in $byAreaTarget){
    $top = $g.Group | Sort-Object -Property delta_med -Descending | Select-Object -First 1
    if ($top) { $pick += $top }
  }
  $bottleneckRows = $pick
}

# ===== HTML 生成 =====
# データをJSONに（PS5.1のConvertTo-JsonはDepth浅いので注意）
$summaryJson    = $summaryRows    | ConvertTo-Json -Depth 5
$bottleneckJson = $bottleneckRows | ConvertTo-Json -Depth 5

$html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<title>NetQuality Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif; margin: 16px; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  .filters { display:flex; gap:8px; flex-wrap: wrap; margin: 8px 0 16px; }
  select, input { padding:6px; border:1px solid #ccc; border-radius: 8px; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 24px; }
  th, td { border-bottom: 1px solid #eee; padding: 8px; text-align: left; }
  th { background: #fafafa; position: sticky; top:0; z-index: 1; }
  .badge { display:inline-block; padding: 2px 8px; border-radius: 999px; background:#eee; font-size: 12px; }
  .rt-ok { background:#e7f7e7; }
  .rt-warn { background:#fff5e0; }
  .rt-bad { background:#fdecec; }
  .loss-bad { background:#fdecec; }
  .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
  .hint { color:#666; font-size:12px; }
  .legend { font-size:12px; color:#444; margin:4px 0 12px; }
  .legend span { margin-right:12px; padding:2px 6px; border-radius:6px; }
</style>
</head>
<body>
  <h1>ネットワーク品質レポート</h1>
  <div class="hint">このレポートは既存CSVを集計して生成されています（HTML単体、外部ライブラリ不要）。</div>

  <h2>概要（エリア → AP → 対象）</h2>
  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="hostSearch" placeholder="対象ホストを検索 (部分一致)" />
    <select id="roleSel">
      <option value="">(すべての役割)</option>
      <option>SaaS</option><option>RouterLAN</option><option>L2</option><option>Internet</option><option>Uncategorized</option><option>Unknown</option>
    </select>
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
        <th>試行数</th>
        <th>RTT中央値</th>
        <th>RTT P95</th>
        <th>損失率(平均)</th>
        <th>MOS(中央値)</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>

  <h2>ボトルネック候補（Hop差分の中央値が最大）</h2>
  <table id="bnTbl">
    <thead>
      <tr>
        <th>エリア</th>
        <th>対象</th>
        <th>候補Hop</th>
        <th>役割</th>
        <th>差分中央値(ms)</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>

<script>
  // ====== データ埋め込み ======
  var summaryRows = $summaryJson;
  var bottleneckRows = $bottleneckJson;

  // ====== UI初期化 ======
  function uniq(vals){ return Array.from(new Set(vals.filter(Boolean))).sort(); }
  function fillSel(el, opts){
    opts.forEach(o => { var op = document.createElement('option'); op.textContent = o; op.value = o; el.appendChild(op); });
  }
  var areaSel = document.getElementById('areaSel');
  var apSel   = document.getElementById('apSel');
  var roleSel = document.getElementById('roleSel');
  var hostSearch = document.getElementById('hostSearch');

  fillSel(areaSel, uniq(summaryRows.map(r=>r.area)));
  fillSel(apSel, uniq(summaryRows.map(r=>r.ap_name).filter(Boolean)));

  [areaSel, apSel, roleSel, hostSearch].forEach(el => el.addEventListener('input', render));

  function colorRtt(td, val){
    if (val == null) return;
    if (val < 50) td.classList.add('rt-ok');
    else if (val < 100) td.classList.add('rt-warn');
    else td.classList.add('rt-bad');
  }
  function colorLoss(td, val){
    if (val == null) return;
    if (val > 3) td.classList.add('loss-bad');
  }

  function render(){
    var area = areaSel.value || "";
    var ap   = apSel.value || "";
    var role = roleSel.value || "";
    var q    = (hostSearch.value || "").toLowerCase();

    // Summary table
    var tbody = document.querySelector('#sumTbl tbody');
    tbody.innerHTML = '';
    summaryRows.forEach(r=>{
      if (area && r.area !== area) return;
      if (ap && (r.ap_name||"") !== ap) return;
      if (role && r.role !== role) return;
      if (q && String(r.host).toLowerCase().indexOf(q) === -1) return;

      var tr = document.createElement('tr');
      function td(t){ var e=document.createElement('td'); e.textContent = (t==null?"":t); return e; }

      tr.appendChild(td(r.area));
      tr.appendChild(td(r.ap_name||""));
      var thost = td(r.host); thost.classList.add('mono'); tr.appendChild(thost);
      tr.appendChild(td(r.role));
      tr.appendChild(td(r.count));
      var rttm = td(r.rtt_med); colorRtt(rttm, r.rtt_med); tr.appendChild(rttm);
      var rttp = td(r.rtt_p95); colorRtt(rttp, r.rtt_p95); tr.appendChild(rttp);
      var loss = td(r.loss_avg); colorLoss(loss, r.loss_avg); tr.appendChild(loss);
      tr.appendChild(td(r.mos_med));
      tbody.appendChild(tr);
    });

    // Bottleneck table
    var btbody = document.querySelector('#bnTbl tbody');
    btbody.innerHTML = '';
    bottleneckRows.forEach(r=>{
      if (area && r.area !== area) return;
      var tr = document.createElement('tr');
      function td(t){ var e=document.createElement('td'); e.textContent = (t==null?"":t); return e; }
      tr.appendChild(td(r.area));
      var tgt = td(r.target); tgt.classList.add('mono'); tr.appendChild(tgt);
      var hop = td((r.node_label||r.hop_ip)); hop.classList.add('mono'); tr.appendChild(hop);
      tr.appendChild(td(r.role||""));
      tr.appendChild(td(r.delta_med));
      btbody.appendChild(tr);
    });
  }

  render();
</script>
</body>
</html>
"@

# 出力
$html | Out-File -FilePath $OutHtml -Encoding utf8
Write-Host "HTMLレポートを出力しました: $OutHtml"
